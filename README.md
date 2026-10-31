# PubMed Article Collector

A simple Python desktop app (Tkinter) that extracts bibliographic metadata from
individual PubMed article webpages and saves it to an Excel file.

It reads the public article page directly using `requests` and `BeautifulSoup`.
It does **not** use the PubMed API or NCBI E-utilities, and needs no paid
services or API keys.
![Uploading image.png…]()


## Extracted fields

`Article_ID | PMID | Article_Title | Abstract | Author_Keywords | DOI | MeSH_Terms | Publication_Year`

- Author keywords and MeSH terms are extracted separately.
- Missing fields are left empty; nothing is generated.
- Multiple keywords / MeSH terms are separated by semicolons.
- Duplicates are blocked by PMID or DOI.
- Records are appended to `Bioinformatics_Articles.xlsx` (created in the app
  folder) and reloaded when the app restarts.

## Requirements

- Python 3.11+ (Tkinter is included with the standard Windows installer)

## Installation (Windows PowerShell)

```powershell
git clone https://github.com/<your-username>/<repo-name>.git
cd <repo-name>
py -3.11 -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
python pubmed_collector.py
```

If PowerShell blocks activation, run once:
`Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`

## Usage

1. Paste a URL such as `https://pubmed.ncbi.nlm.nih.gov/38758314/`.
2. Click **Add Article** (or press Enter).
3. Repeat for more articles; double-click a row to view the full record.
4. Click **Save to Excel**.

## Responsible use

This tool is intended for adding articles one at a time by hand. Do not modify
it to crawl PubMed in bulk. For large-scale retrieval, NCBI provides official
tools and has usage policies you should follow.

## Limitations

- PubMed may change its page layout, which could break extraction.
- Only `pubmed.ncbi.nlm.nih.gov/<PMID>/` URLs are supported.

## License

Add a license of your choice (for example MIT) before publishing.
