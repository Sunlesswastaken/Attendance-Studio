"""Conservative, text-only PDF import. Never silently skip an unreadable page."""
import io
import re

import pandas as pd
import pdfplumber


def pdf_table(data, aliases, key, limit):
    headers = None
    output = []
    with pdfplumber.open(io.BytesIO(data)) as document:
        if not 1 <= len(document.pages) <= 50:
            raise ValueError('Use a PDF with 1–50 pages.')
        for page_no, page in enumerate(document.pages, 1):
            text = page.extract_text() or ''
            if not text.strip():
                raise ValueError(f'PDF page {page_no} has no readable text. Scanned/image PDFs need OCR first; export to searchable PDF or Excel.')
            tables = page.extract_tables()
            if tables:
                blocks = tables
            else:
                # Normalize multi-word header labels only; data retains its original IDs.
                lines = text.splitlines()
                blocks = [[]]
                for line in lines:
                    if re.fullmatch(r'\s*Page\s+\d+(?:\s+of\s+\d+)?\s*', line, re.I):
                        continue
                    if re.search(r'\b(roll|student\s*id|reg(?:istration)?\s*(?:no|number))\b', line, re.I):
                        line = re.sub(r'\b(roll\s*(?:no\.?|number)|student\s+id|registration\s+number|reg\s+no\.?)\b', 'Roll_Number', line, flags=re.I)
                        line = re.sub(r'\bbatch\s+code\b', 'Batch_Code', line, flags=re.I)
                        line = re.sub(r'\b(?:attendance\s+status|(?:cdc|class)\s+attendance)\b', 'Attendance', line, flags=re.I)
                    blocks[0].append(re.split(r'\s+', line.strip()))
            page_rows = 0
            page_header = False
            for block in blocks:
                for raw in block:
                    row = [re.sub(r'\s+', ' ', str(v or '')).strip() for v in raw]
                    if not any(row):
                        continue
                    is_header = any(key(v) in aliases['roll'] for v in row)
                    if is_header:
                        if headers is not None and [key(v) for v in row] != [key(v) for v in headers]:
                            raise ValueError(f'PDF page {page_no} has different columns. Export one consistent attendance table to Excel.')
                        headers = headers or row
                        page_header = True
                        continue
                    if headers is None:
                        continue  # A title before the first header is not a student row.
                    if len(row) != len(headers):
                        raise ValueError(f'PDF page {page_no} has an ambiguous row. Use a clean attendance table or export it to Excel/CSV.')
                    output.append(row)
                    page_rows += 1
                    if len(output) > limit:
                        raise ValueError(f'Maximum {limit:,} student rows per file.')
            if not page_rows and not page_header:
                raise ValueError(f'No attendance rows could be read on PDF page {page_no}. Export a clean text-based table to Excel/CSV.')
    if headers is None:
        raise ValueError('No roll-number header found in the PDF. Include a “Roll Number” column or heading.')
    return {'PDF (all pages)': pd.DataFrame([headers] + output).fillna('')}
