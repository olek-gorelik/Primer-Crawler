import argparse
import json
import os
import re
import sys
import time
import zipfile
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

import requests


ESEARCH_URL = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"
EFETCH_URL = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi"
DEFAULT_QUERY = 'IL11 human (stomach OR gastric) (PCR OR qPCR) (primer OR "forward primer" OR "reverse primer" OR sequence)'
DEFAULT_GENE = "IL11"
RETMAX = 200
TIMEOUT = 15
HEADERS = {"User-Agent": "pmc-primer-crawler/1.0"}
PRIMER_PATTERN = re.compile(r"[ATCGatcg]{18,35}")
SUCCESS_KEYWORDS = [
    "upregulated",
    "downregulated",
    "overexpressed",
    "overexpression",
    "suppressed",
    "suppression",
    "decreased",
    "increased",
    "elevated",
    "reduced",
    "knockdown",
    "silenced",
    "activation",
    "activated",
    "inhibited",
    "inhibition",
    "expression",
]
SUCCESS_CONTEXT_WINDOW = 160
PRIMER_SEARCH_SPAN = 180  # how far after a gene mention to look for primer sequences
DEFAULT_ARTICLE_LIMIT = 200
DEFAULT_EXCEL_PATH = "primers.xlsx"


def log(message):
    """Lightweight logger with timestamp for progress visibility."""
    timestamp = time.strftime("%H:%M:%S")
    print(f"[{timestamp}] {message}", file=sys.stderr, flush=True)


def search_pmc(query, retstart=0, retmax=RETMAX):
    """Search PMC for the provided query and return a list of PMCID strings."""
    params = {
        "db": "pmc",
        "retmax": str(retmax),
        "retstart": str(retstart),
        "retmode": "xml",
        "term": query,
    }
    log(f"Searching PMC for query: {query!r} (start={retstart}, size={retmax})")
    try:
        response = requests.get(ESEARCH_URL, params=params, headers=HEADERS, timeout=TIMEOUT)
        response.raise_for_status()
    except requests.RequestException as exc:
        print(f"ERROR: search failed for query '{query}': {exc}", file=sys.stderr)
        return []

    try:
        root = ET.fromstring(response.text)
    except ET.ParseError as exc:
        print(f"ERROR: could not parse search XML: {exc}", file=sys.stderr)
        return []

    pmc_ids = []
    for id_node in root.findall(".//IdList/Id"):
        raw_id = (id_node.text or "").strip()
        if not raw_id:
            continue
        pmcid = raw_id if raw_id.startswith("PMC") else f"PMC{raw_id}"
        pmc_ids.append(pmcid)

    log(f"Found {len(pmc_ids)} PMC IDs")
    return pmc_ids


def fetch_article_xml(pmcid):
    """Fetch the full article XML for a given PMCID."""
    params = {"db": "pmc", "id": pmcid, "retmode": "xml"}
    log(f"Fetching XML for {pmcid}")
    try:
        response = requests.get(EFETCH_URL, params=params, headers=HEADERS, timeout=TIMEOUT)
        response.raise_for_status()
    except requests.RequestException as exc:
        print(f"ERROR: fetch failed for {pmcid}: {exc}", file=sys.stderr)
        return None

    try:
        return ET.fromstring(response.text)
    except ET.ParseError as exc:
        print(f"ERROR: could not parse XML for {pmcid}: {exc}", file=sys.stderr)
        return None


def _body_without_references(text):
    """Return the article text up to the References section to avoid citation-only hits."""
    lower = text.lower()
    ref_pos = lower.find("references")
    return text if ref_pos == -1 else text[:ref_pos]


def make_gene_pattern(gene_name):
    """Compile a regex that matches the target gene name (IL11 has a richer pattern)."""
    gene = (gene_name or "").strip()
    if not gene:
        gene = DEFAULT_GENE
    if gene.upper() == "IL11":
        pattern = r"\b(?:IL-?11|interleukin[- ]?11)\b"
    else:
        pattern = rf"\b{re.escape(gene)}\b"
    return re.compile(pattern, re.IGNORECASE)


def extract_gene_primers(text, gene_pattern):
    """Extract primer sequences that appear shortly after the target gene is mentioned."""
    if not text:
        return []

    text_lower = text.lower()
    gene_hits = list(gene_pattern.finditer(text_lower))
    if not gene_hits:
        return []

    primers = []
    for gene_hit in gene_hits:
        window_start = gene_hit.end()
        window_end = min(len(text), window_start + PRIMER_SEARCH_SPAN)
        window_slice = text[window_start:window_end]

        found_for_gene = 0
        for primer_match in PRIMER_PATTERN.finditer(window_slice):
            cleaned = "".join(ch for ch in primer_match.group(0).upper() if ch in "ATCG")
            if cleaned:
                primers.append(cleaned)
                found_for_gene += 1
            if found_for_gene >= 2:
                break  # stop after a forward/reverse pair per gene mention

    deduped = []
    seen = set()
    for primer in primers:
        if primer in seen:
            continue
        seen.add(primer)
        deduped.append(primer)
    return deduped


def has_gene_success_evidence(text, gene_pattern):
    """Check for success-related language near the target gene mentions."""
    if not text:
        return False

    text_lower = text.lower()
    for gene_hit in gene_pattern.finditer(text_lower):
        start, end = gene_hit.span()
        window_start = max(0, start - SUCCESS_CONTEXT_WINDOW)
        window_end = min(len(text_lower), end + SUCCESS_CONTEXT_WINDOW)
        window_slice = text_lower[window_start:window_end]
        if any(keyword in window_slice for keyword in SUCCESS_KEYWORDS):
            return True
    return False


def _excel_column_name(col_idx):
    """Convert zero-based column index to Excel column letters (A, B, ... AA...)."""
    name = ""
    idx = col_idx
    while idx >= 0:
        idx, remainder = divmod(idx, 26)
        name = chr(65 + remainder) + name
        idx -= 1
    return name


def _row_xml(row_idx, values):
    """Render a single row of inline string cells for the XLSX sheet."""
    cells = []
    for col_idx, value in enumerate(values):
        ref = f"{_excel_column_name(col_idx)}{row_idx}"
        safe_value = escape(str(value)) if value is not None else ""
        cells.append(f'<c r="{ref}" t="inlineStr"><is><t>{safe_value}</t></is></c>')
    return f'<row r="{row_idx}">' + "".join(cells) + "</row>"


def write_xlsx_table(headers, rows, output_path):
    """Write a minimal XLSX file with the provided headers and row data."""
    sheet_rows = [_row_xml(1, headers)]
    for idx, row in enumerate(rows, start=2):
        sheet_rows.append(_row_xml(idx, row))

    content_types_xml = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>
""".strip()

    rels_xml = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
""".strip()

    workbook_xml = """<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Primers" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
""".strip()

    workbook_rels_xml = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
""".strip()

    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        "<sheetData>"
        + "".join(sheet_rows)
        + "</sheetData></worksheet>"
    )

    output_abs_path = os.path.abspath(output_path)
    with zipfile.ZipFile(output_abs_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml)
        zf.writestr("_rels/.rels", rels_xml)
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)

    return output_abs_path


def resolve_output_path(path, allow_overwrite=False):
    """Return a file path that will not clobber existing files unless allowed."""
    if allow_overwrite or not os.path.exists(path):
        return path

    base, ext = os.path.splitext(path)
    counter = 1
    while True:
        candidate = f"{base}_{counter}{ext}"
        if not os.path.exists(candidate):
            return candidate
        counter += 1


def build_primer_rows(results, gene_label):
    """Transform crawler results into a 4-column row list for Excel export."""
    rows = []
    for record in results:
        primers = record.get("primers") or []
        if not primers:
            continue
        url = record.get("url", "")
        for idx in range(0, len(primers), 2):
            forward = primers[idx]
            reverse = primers[idx + 1] if idx + 1 < len(primers) else ""
            rows.append([gene_label, url, forward, reverse])
    return rows


def infer_gene_label(query, fallback=DEFAULT_GENE):
    """Pull a simple gene label from the query for first-column tagging."""
    first_token = (query or "").strip().split()[0:1]
    return first_token[0] if first_token else fallback


def _extract_article_text(xml_root):
    """Concatenate all text nodes from an article XML tree."""
    return " ".join(xml_root.itertext())


def crawl(query, gene_pattern, gene_label, article_limit=DEFAULT_ARTICLE_LIMIT, retstart=0, retmax=RETMAX):
    """Perform search, fetch articles, and extract primer data for each PMCID."""
    results = []
    pmc_ids = search_pmc(query, retstart=retstart, retmax=retmax)
    pmc_ids = pmc_ids[:article_limit]
    log(f"Processing {len(pmc_ids)} articles from offset {retstart}")

    for pmcid in pmc_ids:
        primers = []
        success_evidence = False
        xml_root = fetch_article_xml(pmcid)
        if xml_root is not None:
            article_text = _extract_article_text(xml_root)
            body_text = _body_without_references(article_text)

            if not gene_pattern.search(body_text):
                log(f"{pmcid}: skipping (no {gene_label} mention outside references)")
                continue

            primers = extract_gene_primers(body_text, gene_pattern)
            success_evidence = has_gene_success_evidence(body_text, gene_pattern)

            log(
                f"{pmcid}: extracted {len(primers)} {gene_label}-linked primer sequences; "
                f"success evidence={success_evidence}"
            )

        results.append(
            {
                "pmcid": pmcid,
                "url": f"https://pmc.ncbi.nlm.nih.gov/articles/{pmcid}/",
                "has_primers": bool(primers),
                "primers": primers,
                "success_evidence": success_evidence,
            }
        )

    return results


def parse_args():
    """CLI argument parser for query override and export options."""
    parser = argparse.ArgumentParser(
        description="Find gene-specific primers in PubMed Central (default: IL11) and export results."
    )
    parser.add_argument(
        "query",
        nargs="*",
        help="Optional override for the search query. Example: EGR1 human (PCR OR qPCR) primer",
    )
    parser.add_argument(
        "-n",
        "--article-limit",
        type=int,
        default=DEFAULT_ARTICLE_LIMIT,
        help=f"Number of PMC articles to process (default: {DEFAULT_ARTICLE_LIMIT})",
    )
    parser.add_argument(
        "--page",
        type=int,
        default=0,
        help="Zero-based page of results to fetch (multiplies page-size for the start offset).",
    )
    parser.add_argument(
        "--page-size",
        type=int,
        default=RETMAX,
        help=f"Number of PMC IDs to request per page (default: {RETMAX}).",
    )
    parser.add_argument(
        "-x",
        "--excel",
        default=DEFAULT_EXCEL_PATH,
        help=f"Path for the Excel table (.xlsx). Default: {DEFAULT_EXCEL_PATH}",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Allow overwriting the Excel file (otherwise a _1, _2 suffix is added).",
    )
    parser.add_argument(
        "-g",
        "--gene",
        help="Gene label to use in the first Excel column (defaults to the first token in the query).",
    )
    parser.add_argument(
        "-t",
        "--target-gene",
        help="Gene name to search around in the article text (default: IL11).",
    )
    parser.add_argument(
        "--skip-json",
        action="store_true",
        help="Suppress printing the raw JSON crawl data to stdout.",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    override_query = " ".join(arg for arg in args.query if arg.strip())
    query = override_query or DEFAULT_QUERY
    target_gene = args.target_gene or infer_gene_label(query, fallback=DEFAULT_GENE)
    gene_label = args.gene or target_gene
    gene_pattern = make_gene_pattern(target_gene)
    retstart = max(0, args.page * args.page_size)
    excel_target = resolve_output_path(args.excel, allow_overwrite=args.overwrite)

    log("Starting crawl")
    data = crawl(
        query,
        gene_pattern,
        gene_label,
        article_limit=args.article_limit,
        retstart=retstart,
        retmax=args.page_size,
    )
    log(f"Completed crawl; {len(data)} records")

    primer_rows = build_primer_rows(data, gene_label)
    if primer_rows:
        headers = ["Gene", "URL", "Primer 1", "Primer 2"]
        excel_path = write_xlsx_table(headers, primer_rows, excel_target)
        log(f"Wrote Excel table ({len(primer_rows)} row(s)) to {excel_path}")
    else:
        log("No primer sequences found; Excel export skipped")

    if not args.skip_json:
        print(json.dumps(data, indent=2))
