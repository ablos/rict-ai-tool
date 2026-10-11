"""Turn a PDF into page-numbered markdown with Mistral OCR, cached per PDF.

The OCR result for a given PDF never changes, so it is stored once under
cache/ocr/<sha256 of the PDF>.json and reused on every later run.
"""

import base64
import hashlib
import json
from pathlib import Path
from mistralai.client import Mistral

OCR_MODEL = "mistral-ocr-4-1"
CACHE_DIR = Path("cache/ocr")

def file_sha256(path: str | Path) -> str:
    """Fingerprint of a file's contents: same bytes, same hash, whatever the file name."""
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()

def ocr_pdf(client: Mistral, pdf_path: str | Path) -> dict:
    """Return the OCR text of a PDF, from the cache if it has been processed before.

    Pages are numbered from 1 as in a PDF viewer, which can differ from the
    page numbers printed in a journal.
    """

    pdf_path = Path(pdf_path)
    sha256 = file_sha256(pdf_path)
    cache_file = CACHE_DIR / f"{sha256}.json"
    
    if cache_file.exists():
        return json.loads(cache_file.read_text(encoding='utf-8'))

    encoded = base64.b64encode(pdf_path.read_bytes()).decode('ascii')
    response = client.ocr.process(
        model=OCR_MODEL,
        document={"type": "document_url", "document_url": f"data:application/pdf;base64,{encoded}"},
        table_format="markdown",
        extract_header=True, # keep running heads and page footers out of the body text
        extract_footer=True,
        include_blocks=False, # layout boxes, not needed
    )

    result = {
        "pdf": pdf_path.name,
        "sha256": sha256,
        "model": response.model,
        "pages": [
            {
                "page": page.index + 1,
                "markdown": _inline_tables(page),
                "header": page.header or None,
                "footer": page.footer or None,
            }

            for page in response.pages
        ],
    }

    # A damaged PDF can come back as empty pages; stop here instead of caching nothing.
    if not any(page["markdown"].strip() for page in result["pages"]):
        raise ValueError(f"OCR found no text in {pdf_path.name}; the PDF may be damaged")

    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    cache_file.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    return result

def _inline_tables(page) -> str:
    """Put tables back into the page text, where OCR left a [tbl-0.md](tbl-0.md) link."""
    markdown = page.markdown
    
    for table in page.tables or []:
        markdown = markdown.replace(f"[{table.id}]({table.id})", table.content)

    return markdown

if __name__ == "__main__":
    import os
    import sys
    from dotenv import load_dotenv

    load_dotenv()
    client = Mistral(api_key=os.environ["MISTRAL_API_KEY"])
    ocr = ocr_pdf(client, sys.argv[1])
    print(f"{ocr['pdf']}: {len(ocr['pages'])} pages, OCR model {ocr['model']}\n")
    print(ocr["pages"][0]["markdown"][:1500])