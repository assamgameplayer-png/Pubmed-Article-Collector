"""
PubMed Article Collector (no API)
---------------------------------
Paste one PubMed article URL at a time, extract bibliographic metadata directly
from the article webpage (requests + BeautifulSoup), and save it to
Bioinformatics_Articles.xlsx (openpyxl) in the same folder as this script.

Python 3.11 | Tkinter | requests | beautifulsoup4 | openpyxl
"""

import re
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------------
# Settings
# ----------------------------------------------------------------------------
APP_FOLDER = Path(__file__).resolve().parent
EXCEL_FILE = APP_FOLDER / "Bioinformatics_Articles.xlsx"

COLUMNS = [
    "Article_ID", "PMID", "Article_Title", "Abstract",
    "Author_Keywords", "DOI", "MeSH_Terms", "Publication_Year",
]

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml",
    "Accept-Language": "en-US,en;q=0.9",
}

URL_PATTERN = re.compile(
    r"^https?://pubmed\.ncbi\.nlm\.nih\.gov/(\d+)/?(?:[?#].*)?$", re.IGNORECASE
)

# Characters Excel/openpyxl cannot store
ILLEGAL_XML_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------
def clean_text(text: str) -> str:
    """Collapse whitespace (including newlines) into single spaces."""
    return re.sub(r"\s+", " ", text or "").strip()


def excel_safe(value):
    """Remove characters that would crash openpyxl; keep everything else."""
    if isinstance(value, str):
        return ILLEGAL_XML_CHARS.sub("", value)
    return value


def normalize_doi(doi: str) -> str:
    """Lower-case, prefix-free DOI used only for duplicate comparison."""
    doi = (doi or "").strip().lower()
    doi = re.sub(r"^https?://(dx\.)?doi\.org/", "", doi)
    return doi


# ----------------------------------------------------------------------------
# Web page extraction
# ----------------------------------------------------------------------------
def fetch_html(url: str) -> str:
    try:
        response = requests.get(url, headers=HEADERS, timeout=25)
    except requests.exceptions.Timeout:
        raise RuntimeError("The request timed out. Check your internet and try again.")
    except requests.exceptions.ConnectionError:
        raise RuntimeError("Could not connect. Check your internet connection.")
    except requests.exceptions.RequestException as exc:
        raise RuntimeError(f"Could not access the webpage: {exc}")

    if response.status_code == 404:
        raise RuntimeError("Article not found (HTTP 404). Check the URL.")
    if response.status_code == 429:
        raise RuntimeError("PubMed asked to slow down (HTTP 429). Wait a minute and retry.")
    if response.status_code != 200:
        raise RuntimeError(f"Webpage returned HTTP {response.status_code}.")
    return response.text


def _meta(soup: BeautifulSoup, name: str) -> str:
    tag = soup.find("meta", attrs={"name": name})
    return clean_text(tag["content"]) if tag and tag.get("content") else ""


def extract_pmid(soup: BeautifulSoup, url_pmid: str) -> str:
    tag = soup.select_one("strong.current-id")
    if tag and clean_text(tag.get_text()).isdigit():
        return clean_text(tag.get_text())
    for name in ("citation_pmid", "log_displayeduids"):
        value = _meta(soup, name)
        if value.isdigit():
            return value
    return url_pmid


def extract_title(soup: BeautifulSoup) -> str:
    tag = soup.select_one("h1.heading-title")
    if tag:
        return clean_text(tag.get_text(" "))
    return _meta(soup, "citation_title")


def _abstract_container(soup: BeautifulSoup):
    return (
        soup.select_one("div.abstract#eng-abstract")
        or soup.select_one("div#eng-abstract")
        or soup.select_one("div#abstract")
        or soup.select_one("div.abstract")
    )


def extract_abstract(soup: BeautifulSoup) -> str:
    container = _abstract_container(soup)
    if not container:
        return ""
    content = container.select_one("div.abstract-content") or container

    parts = []
    for p in content.find_all("p"):
        label_tag = p.find("strong", class_="sub-title")
        label = clean_text(label_tag.get_text()) if label_tag else ""
        # The author-keywords paragraph also lives here; it is not abstract text.
        if label.lower().rstrip(":").strip() in ("keywords", "keyword"):
            continue
        if label_tag:
            label_tag.extract()  # remove so the label is not repeated in the body
        body = clean_text(p.get_text(" "))
        if not body:
            continue
        parts.append(f"{label} {body}".strip() if label else body)
    return "\n".join(parts)


def extract_author_keywords(soup: BeautifulSoup) -> str:
    """Author keywords appear as a 'Keywords:' paragraph under the abstract.
    They are NOT the MeSH terms (those live in the separate MeSH section)."""
    # Search the outer abstract block first, then the whole page as a fallback.
    scopes = [soup.select_one("div.abstract"), soup.select_one("div#abstract"), soup]
    container = next((c for c in scopes if c is not None and any(
        clean_text(s.get_text()).lower().rstrip(":").strip() in ("keywords", "keyword")
        for s in c.find_all("strong", class_="sub-title"))), None)
    if container is None:
        return ""
    for strong in container.find_all("strong", class_="sub-title"):
        label = clean_text(strong.get_text()).lower().rstrip(":").strip()
        if label in ("keywords", "keyword"):
            p = strong.find_parent("p")
            if not p:
                continue
            strong.extract()
            raw = clean_text(p.get_text(" "))
            items = [clean_text(k) for k in raw.split(";")]
            items = [k for k in items if k]
            if items:
                # PubMed ends the list with a full stop; it is display punctuation.
                items[-1] = items[-1].rstrip(".").strip()
            return "; ".join(k for k in items if k)
    return ""


def extract_mesh_terms(soup: BeautifulSoup) -> str:
    section = soup.select_one("div#mesh-terms") or soup.select_one("div.mesh-terms")
    if not section:
        return ""
    terms = []
    for li in section.select("ul li"):
        trigger = li.select_one("button.keyword-actions-trigger")
        text = clean_text((trigger or li).get_text(" "))
        text = text.rstrip("*").strip()  # '*' only marks a major topic
        text = re.sub(r"\s+/\s+", "/", text)
        if text and text not in terms:
            terms.append(text)
    return "; ".join(terms)


def extract_doi(soup: BeautifulSoup) -> str:
    link = soup.select_one("span.identifier.doi a.id-link") or soup.select_one(
        "span.identifier.doi a"
    )
    if link:
        return clean_text(link.get_text())
    for name in ("citation_doi", "dc.identifier"):
        value = _meta(soup, name)
        if value:
            value = re.sub(r"^doi:\s*", "", value, flags=re.IGNORECASE)
            if value.startswith("10."):
                return value
    return ""


def extract_year(soup: BeautifulSoup) -> str:
    for name in ("citation_publication_date", "citation_date", "citation_online_date"):
        match = re.search(r"\b(1[6-9]\d{2}|20\d{2})\b", _meta(soup, name))
        if match:
            return match.group(1)
    cite = soup.select_one("span.cit")
    if cite:
        match = re.search(r"\b(1[6-9]\d{2}|20\d{2})\b", cite.get_text())
        if match:
            return match.group(1)
    return ""


def extract_article(url: str) -> dict:
    """Download one PubMed article page and return its metadata as a dict."""
    url = url.strip()
    match = URL_PATTERN.match(url)
    if not match:
        raise RuntimeError(
            "Invalid URL. Use a link like https://pubmed.ncbi.nlm.nih.gov/38758314/"
        )
    url_pmid = match.group(1)
    page_url = f"https://pubmed.ncbi.nlm.nih.gov/{url_pmid}/"

    soup = BeautifulSoup(fetch_html(page_url), "html.parser")

    title = extract_title(soup)
    if not title:
        raise RuntimeError(
            "Could not find the article title. The page layout may have changed "
            "or PubMed returned a different page."
        )

    return {
        "PMID": extract_pmid(soup, url_pmid),
        "Article_Title": title,
        "Abstract": extract_abstract(soup),
        "Author_Keywords": extract_author_keywords(soup),
        "DOI": extract_doi(soup),
        "MeSH_Terms": extract_mesh_terms(soup),
        "Publication_Year": extract_year(soup),
    }


# ----------------------------------------------------------------------------
# Excel storage
# ----------------------------------------------------------------------------
def load_records_from_excel() -> list[dict]:
    """Read previously saved records so work continues after a restart."""
    if not EXCEL_FILE.exists():
        return []
    wb = load_workbook(EXCEL_FILE)
    ws = wb.active
    records = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or all(v is None for v in row):
            continue
        record = {}
        for col_name, value in zip(COLUMNS, row):
            record[col_name] = "" if value is None else value
        record["Article_ID"] = int(record["Article_ID"]) if str(record["Article_ID"]).isdigit() else 0
        record["PMID"] = str(record["PMID"])
        record["Publication_Year"] = str(record["Publication_Year"])
        records.append(record)
    wb.close()
    return records


def append_records_to_excel(new_records: list[dict]) -> None:
    """Append rows to the workbook (creating it, with headers, if needed)."""
    if EXCEL_FILE.exists():
        wb = load_workbook(EXCEL_FILE)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Articles"
        ws.append(COLUMNS)
        for cell in ws[1]:
            cell.font = Font(bold=True)
        ws.freeze_panes = "A2"
        widths = {"A": 11, "B": 12, "C": 60, "D": 80, "E": 40, "F": 28, "G": 60, "H": 17}
        for letter, width in widths.items():
            ws.column_dimensions[letter].width = width

    for rec in new_records:
        ws.append([excel_safe(rec[c]) for c in COLUMNS])
        row_idx = ws.max_row
        for col_idx in range(1, len(COLUMNS) + 1):
            ws.cell(row=row_idx, column=col_idx).alignment = Alignment(
                vertical="top", wrap_text=False
            )
        # Keep PMID and year as text-like values consistent with the page
        ws.cell(row=row_idx, column=2).number_format = "@"
        ws.cell(row=row_idx, column=8).number_format = "@"

    wb.save(EXCEL_FILE)
    wb.close()


# ----------------------------------------------------------------------------
# GUI
# ----------------------------------------------------------------------------
class PubMedApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PubMed Article Collector")
        self.geometry("1150x600")
        self.minsize(900, 450)

        self.records: list[dict] = []
        self.unsaved: list[dict] = []

        self._build_widgets()
        self._load_existing()

    # ---- layout -----------------------------------------------------------
    def _build_widgets(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")

        ttk.Label(top, text="PubMed article URL:").grid(row=0, column=0, sticky="w")
        self.url_var = tk.StringVar()
        self.url_entry = ttk.Entry(top, textvariable=self.url_var, width=80)
        self.url_entry.grid(row=1, column=0, sticky="ew", padx=(0, 8))
        self.url_entry.bind("<Return>", lambda e: self.add_article())

        self.add_btn = ttk.Button(top, text="Add Article", command=self.add_article)
        self.add_btn.grid(row=1, column=1, padx=4)
        self.save_btn = ttk.Button(top, text="Save to Excel", command=self.save_to_excel)
        self.save_btn.grid(row=1, column=2, padx=4)
        top.columnconfigure(0, weight=1)

        info = ttk.Frame(self, padding=(10, 0))
        info.pack(fill="x")
        self.counter_var = tk.StringVar(value="Total articles collected: 0")
        ttk.Label(info, textvariable=self.counter_var, font=("Segoe UI", 10, "bold")).pack(
            anchor="w"
        )
        self.status_var = tk.StringVar(value="Paste a PubMed URL and click 'Add Article'.")
        self.status_label = tk.Label(
            info, textvariable=self.status_var, anchor="w", justify="left",
            wraplength=1100, fg="black",
        )
        self.status_label.pack(fill="x", pady=(4, 6))

        table_frame = ttk.Frame(self, padding=10)
        table_frame.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(table_frame, columns=COLUMNS, show="headings")
        widths = {
            "Article_ID": 70, "PMID": 85, "Article_Title": 260, "Abstract": 260,
            "Author_Keywords": 160, "DOI": 150, "MeSH_Terms": 200, "Publication_Year": 110,
        }
        for col in COLUMNS:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=widths[col], anchor="w", stretch=False)

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        self.tree.bind("<Double-1>", self.show_details)
        ttk.Label(
            self, text="Tip: double-click a row to read the full record.", padding=(10, 0, 10, 8)
        ).pack(anchor="w")

        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.url_entry.focus_set()

    # ---- helpers ----------------------------------------------------------
    def set_status(self, message: str, ok: bool | None = None):
        color = "black" if ok is None else ("#0a7d28" if ok else "#c0392b")
        self.status_label.config(fg=color)
        self.status_var.set(message)
        self.update_idletasks()

    def update_counter(self):
        pending = len(self.unsaved)
        text = f"Total articles collected: {len(self.records)}"
        if pending:
            text += f"   ({pending} not yet saved to Excel)"
        self.counter_var.set(text)

    def insert_row(self, record: dict):
        values = []
        for col in COLUMNS:
            value = str(record[col])
            if col == "Abstract":
                value = value.replace("\n", " ")
            values.append(value)
        item = self.tree.insert("", "end", values=values)
        self.tree.see(item)

    def _load_existing(self):
        try:
            self.records = load_records_from_excel()
        except Exception as exc:
            messagebox.showerror(
                "Could not read Excel file",
                f"The existing file could not be read:\n{exc}\n\n"
                "Close it in Excel if it is open, then restart the app.",
            )
            self.records = []
        for rec in self.records:
            self.insert_row(rec)
        self.update_counter()
        if self.records:
            self.set_status(
                f"Loaded {len(self.records)} existing article(s) from {EXCEL_FILE.name}.", True
            )

    def _is_duplicate(self, pmid: str, doi: str) -> str:
        doi_norm = normalize_doi(doi)
        for rec in self.records:
            if pmid and str(rec["PMID"]) == pmid:
                return f"PMID {pmid}"
            if doi_norm and normalize_doi(str(rec["DOI"])) == doi_norm:
                return f"DOI {doi}"
        return ""

    def _next_id(self) -> int:
        return max((int(r["Article_ID"]) for r in self.records), default=0) + 1

    # ---- actions ----------------------------------------------------------
    def add_article(self):
        url = self.url_var.get().strip()
        if not url:
            self.set_status("Please paste a PubMed URL first.", False)
            return

        # Quick duplicate check on the PMID in the URL before downloading
        m = URL_PATTERN.match(url)
        if m and self._is_duplicate(m.group(1), ""):
            self.set_status(f"Duplicate skipped: PMID {m.group(1)} is already collected.", False)
            return

        self.add_btn.config(state="disabled")
        self.config(cursor="watch")
        self.set_status("Fetching article from PubMed...", None)
        try:
            data = extract_article(url)
            dup = self._is_duplicate(data["PMID"], data["DOI"])
            if dup:
                self.set_status(f"Duplicate skipped: {dup} is already collected.", False)
                return

            record = {"Article_ID": self._next_id(), **data}
            self.records.append(record)
            self.unsaved.append(record)
            self.insert_row(record)
            self.update_counter()
            self.url_var.set("")

            notes = []
            if not data["Abstract"]:
                notes.append("no abstract")
            if not data["Author_Keywords"]:
                notes.append("no author keywords")
            if not data["MeSH_Terms"]:
                notes.append("no MeSH terms")
            if not data["DOI"]:
                notes.append("no DOI")
            extra = f" (not available: {', '.join(notes)})" if notes else ""
            self.set_status(
                f"Success: added PMID {data['PMID']} as Article_ID {record['Article_ID']}.{extra}",
                True,
            )
        except RuntimeError as exc:
            self.set_status(f"Error: {exc}", False)
        except Exception as exc:  # unexpected parsing problem
            self.set_status(f"Unexpected error while extracting data: {exc}", False)
        finally:
            self.add_btn.config(state="normal")
            self.config(cursor="")
            self.url_entry.focus_set()

    def save_to_excel(self) -> bool:
        if not self.unsaved:
            self.set_status("Nothing new to save. All records are already in the Excel file.", None)
            return True
        try:
            append_records_to_excel(self.unsaved)
        except PermissionError:
            self.set_status(
                "Error: could not save. Close Bioinformatics_Articles.xlsx in Excel and try again.",
                False,
            )
            return False
        except Exception as exc:
            self.set_status(f"Error while saving Excel file: {exc}", False)
            return False
        count = len(self.unsaved)
        self.unsaved.clear()
        self.update_counter()
        self.set_status(f"Saved {count} new article(s) to {EXCEL_FILE}", True)
        return True

    def show_details(self, _event=None):
        selected = self.tree.selection()
        if not selected:
            return
        article_id = self.tree.item(selected[0], "values")[0]
        record = next((r for r in self.records if str(r["Article_ID"]) == str(article_id)), None)
        if not record:
            return
        win = tk.Toplevel(self)
        win.title(f"Article {article_id}")
        win.geometry("800x600")
        text = tk.Text(win, wrap="word", padx=10, pady=10)
        scroll = ttk.Scrollbar(win, command=text.yview)
        text.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        text.pack(fill="both", expand=True)
        for col in COLUMNS:
            text.insert("end", f"{col}:\n", ("bold",))
            text.insert("end", f"{record[col]}\n\n")
        text.tag_configure("bold", font=("Segoe UI", 10, "bold"))
        text.config(state="disabled")

    def on_close(self):
        if self.unsaved:
            answer = messagebox.askyesnocancel(
                "Unsaved articles",
                f"You have {len(self.unsaved)} article(s) not saved to Excel.\n\nSave before closing?",
            )
            if answer is None:
                return
            if answer and not self.save_to_excel():
                return
        self.destroy()


if __name__ == "__main__":
    PubMedApp().mainloop()
