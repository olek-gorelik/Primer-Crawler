import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pmc_primer_crawler as crawler


class PrimerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Primer Crawler")
        self.queue = queue.Queue()
        self.worker = None
        self.running = False
        self.last_rows = []
        self.last_gene = ""

        self._build_ui()
        self._poll_queue()

    def _build_ui(self):
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True)

        self.form_frame = ttk.Frame(notebook, padding=12)
        self.results_frame = ttk.Frame(notebook, padding=12)
        notebook.add(self.form_frame, text="Crawler")
        notebook.add(self.results_frame, text="Results")

        # Form inputs
        grid = ttk.Frame(self.form_frame)
        grid.pack(fill=tk.X, pady=(0, 12))

        self.gene_var = tk.StringVar(value=crawler.DEFAULT_GENE)
        self.query_var = tk.StringVar(value=crawler.DEFAULT_QUERY)
        self.limit_var = tk.StringVar(value=str(crawler.DEFAULT_ARTICLE_LIMIT))
        self.page_var = tk.StringVar(value="0")
        self.page_size_var = tk.StringVar(value=str(crawler.RETMAX))
        self.excel_var = tk.StringVar(value=crawler.DEFAULT_EXCEL_PATH)
        self.create_excel_var = tk.BooleanVar(value=True)

        def add_row(label_text, widget, row):
            ttk.Label(grid, text=label_text).grid(row=row, column=0, sticky=tk.W, padx=(0, 8), pady=4)
            widget.grid(row=row, column=1, sticky=tk.EW, pady=4)

        gene_entry = ttk.Entry(grid, textvariable=self.gene_var, width=40)
        query_entry = ttk.Entry(grid, textvariable=self.query_var, width=60)
        limit_entry = ttk.Entry(grid, textvariable=self.limit_var, width=12)
        page_entry = ttk.Entry(grid, textvariable=self.page_var, width=12)
        page_size_entry = ttk.Entry(grid, textvariable=self.page_size_var, width=12)
        excel_entry = ttk.Entry(grid, textvariable=self.excel_var, width=40)

        add_row("Gene", gene_entry, 0)
        add_row("Query", query_entry, 1)
        add_row("Article limit", limit_entry, 2)
        add_row("Page number (0-based)", page_entry, 3)
        add_row("Page size", page_size_entry, 4)
        add_row("Excel path", excel_entry, 5)

        excel_chk = ttk.Checkbutton(grid, text="Create Excel automatically", variable=self.create_excel_var)
        excel_chk.grid(row=6, column=1, sticky=tk.W, pady=4)

        grid.columnconfigure(1, weight=1)

        # Start button
        btn_frame = ttk.Frame(self.form_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 8))
        self.start_btn = ttk.Button(btn_frame, text="Start", command=self.start_crawl)
        self.start_btn.pack(side=tk.LEFT)

        # Progress log
        ttk.Label(self.form_frame, text="Progress").pack(anchor=tk.W)
        self.progress = tk.Text(self.form_frame, height=12, state=tk.DISABLED, wrap=tk.WORD)
        self.progress.pack(fill=tk.BOTH, expand=True, pady=(4, 0))

        # Results tab
        controls = ttk.Frame(self.results_frame)
        controls.pack(fill=tk.X, pady=(0, 8))
        self.save_btn = ttk.Button(controls, text="Save to Excel", command=self.save_results)
        self.save_btn.pack(side=tk.LEFT)
        self.results_status = ttk.Label(controls, text="No results yet.")
        self.results_status.pack(side=tk.LEFT, padx=(8, 0))

        columns = ("gene", "url", "primer1", "primer2")
        self.tree = ttk.Treeview(self.results_frame, columns=columns, show="headings", height=12)
        for col, heading in zip(columns, ["Gene", "URL", "Primer 1", "Primer 2"]):
            self.tree.heading(col, text=heading)
            self.tree.column(col, width=150 if col != "url" else 320, anchor=tk.W)
        self.tree.pack(fill=tk.BOTH, expand=True)

    def start_crawl(self):
        if self.running:
            return
        try:
            limit = int(self.limit_var.get())
            page = int(self.page_var.get())
            page_size = int(self.page_size_var.get())
        except ValueError:
            messagebox.showerror("Invalid input", "Article limit, page number and page size must be integers.")
            return

        gene_input = self.gene_var.get().strip()
        query = self.query_var.get().strip() or crawler.DEFAULT_QUERY
        target_gene = gene_input or crawler.infer_gene_label(query, fallback=crawler.DEFAULT_GENE)
        gene_label = target_gene
        excel_path = self.excel_var.get().strip() or crawler.DEFAULT_EXCEL_PATH
        create_excel = self.create_excel_var.get()

        self.running = True
        self.start_btn.state(["disabled"])
        self._clear_progress()
        self._log(f"Starting with gene={target_gene}, limit={limit}, page={page}, size={page_size}")

        self.worker = threading.Thread(
            target=self._run_crawl,
            args=(query, target_gene, gene_label, limit, page, page_size, excel_path, create_excel),
            daemon=True,
        )
        self.worker.start()

    def _run_crawl(self, query, target_gene, gene_label, limit, page, page_size, excel_path, create_excel):
        old_log = crawler.log

        def gui_log(msg):
            self.queue.put(("log", msg))

        crawler.log = gui_log
        try:
            retstart = max(0, page * page_size)
            gene_pattern = crawler.make_gene_pattern(target_gene)
            data = crawler.crawl(
                query,
                gene_pattern,
                gene_label,
                article_limit=limit,
                retstart=retstart,
                retmax=page_size,
            )
            rows = crawler.build_primer_rows(data, gene_label)
            excel_written = None
            if create_excel and rows:
                resolved = crawler.resolve_output_path(excel_path, allow_overwrite=False)
                excel_written = crawler.write_xlsx_table(["Gene", "URL", "Primer 1", "Primer 2"], rows, resolved)
                gui_log(f"Excel saved to {excel_written}")
            self.queue.put(("done", {"rows": rows, "gene": gene_label, "excel_path": excel_written}))
        except Exception as exc:  # pylint: disable=broad-except
            self.queue.put(("error", str(exc)))
        finally:
            crawler.log = old_log

    def _poll_queue(self):
        try:
            while True:
                kind, payload = self.queue.get_nowait()
                if kind == "log":
                    self._log(payload)
                elif kind == "done":
                    self._handle_done(payload)
                elif kind == "error":
                    self._handle_error(payload)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    def _handle_done(self, payload):
        self.running = False
        self.start_btn.state(["!disabled"])
        self.last_rows = payload.get("rows", [])
        self.last_gene = payload.get("gene", "")
        excel_path = payload.get("excel_path")

        self._populate_results(self.last_rows)
        msg = f"Completed. Rows: {len(self.last_rows)}"
        if excel_path:
            msg += f"; Excel: {excel_path}"
        self.results_status.config(text=msg)
        self._log(msg)

    def _handle_error(self, message):
        self.running = False
        self.start_btn.state(["!disabled"])
        self._log(f"ERROR: {message}")
        messagebox.showerror("Error", message)

    def _populate_results(self, rows):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for row in rows:
            self.tree.insert("", tk.END, values=row)

    def save_results(self):
        if not self.last_rows:
            messagebox.showinfo("No results", "There are no results to save.")
            return
        path = filedialog.asksaveasfilename(
            title="Save Excel file",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile="primers.xlsx",
        )
        if not path:
            return
        path = os.path.abspath(path)
        crawler.write_xlsx_table(["Gene", "URL", "Primer 1", "Primer 2"], self.last_rows, path)
        self._log(f"Saved Excel to {path}")
        messagebox.showinfo("Saved", f"Excel saved to:\n{path}")

    def _clear_progress(self):
        self.progress.configure(state=tk.NORMAL)
        self.progress.delete("1.0", tk.END)
        self.progress.configure(state=tk.DISABLED)

    def _log(self, message):
        self.progress.configure(state=tk.NORMAL)
        self.progress.insert(tk.END, f"{message}\n")
        self.progress.see(tk.END)
        self.progress.configure(state=tk.DISABLED)


def main():
    root = tk.Tk()
    PrimerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
