# Primer Crawler

Python utility that searches PubMed Central (PMC) articles for gene-specific PCR primers, flags success-related context, and exports the results to a simple Excel file. A Tkinter GUI is included for non-CLI users.

## Features
- Searches PMC via NCBI E-utilities with a user-provided query and paging controls.
- Extracts forward/reverse primer sequences that appear near mentions of the target gene.
- Marks articles that include success/efficacy language around the gene.
- Exports primers to a minimal `.xlsx` table without external spreadsheet libraries.
- Tkinter GUI for point-and-click use; CLI for scripting and reproducible runs.

## Requirements
- Python 3.9+ (tested with Python 3.12 for the GUI).
- `requests` library (`pip install requests`).
- Internet access to reach `eutils.ncbi.nlm.nih.gov`.
- Tkinter (bundled with most Python installations) for the GUI.

## Installation
Clone or download this repository, then install the single dependency:

```bash
pip install requests
```

On macOS, you can double-click `Run Primer Crawler.command` (uses the system Python 3.12) to start the GUI. On other platforms, launch the GUI with `python primer_gui.py`.

## Usage (CLI)
Run the crawler directly to fetch PMC articles, extract primers, and write an Excel table:

```bash
python pmc_primer_crawler.py "IL11 human (stomach OR gastric) (PCR OR qPCR) primer"
```

Common flags:
- `-n, --article-limit`: number of PMC articles to process (default: 200).
- `--page` and `--page-size`: paging controls for PMC search results.
- `-g, --gene`: label to place in the first Excel column (defaults to the first token in the query).
- `-t, --target-gene`: gene name to search around in the article text (default: IL11).
- `-x, --excel`: output Excel path (default: `primers.xlsx`); add `--overwrite` to allow replacing existing files.
- `--skip-json`: suppress raw JSON output to stdout.

## Usage (GUI)
1. Launch `Run Primer Crawler.command` on macOS, or run `python primer_gui.py`.
2. Enter the gene/query, article limit, page number/size, and Excel output path.
3. Click **Start** to crawl; progress logs show in the Crawler tab.
4. View extracted primers in the Results tab and click **Save to Excel** if needed.

## Output
- Excel file with columns: `Gene | URL | Primer 1 | Primer 2`.
- JSON crawl data printed to stdout in CLI mode (unless `--skip-json` is used), containing PMCID, URL, extracted primers, and whether success evidence was detected.

## Notes
- Be considerate of NCBI rate limits; the code uses simple HTTP requests without retries/backoff.
- The crawler stops after the requested number of PMC IDs, even if no primers are found.
- IL11 has a richer matching pattern (`IL-11`, `interleukin-11`); other genes match exact names.
