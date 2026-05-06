import json
import sys
from pathlib import Path

import pandas as pd


def clean_frame(frame):
    frame = frame.dropna(how="all")
    frame = frame.loc[:, ~frame.columns.astype(str).str.startswith("Unnamed")]
    frame.columns = [str(column).strip() for column in frame.columns]
    return frame.fillna("").astype(str)


def frame_records(frame, source):
    cleaned = clean_frame(frame)
    records = cleaned.to_dict(orient="records")
    for record in records:
        for key, value in list(record.items()):
            text = str(value).strip()
            record[key] = text[:240]
        record["_source"] = source
    return records


def read_pdf(path):
    try:
        import pdfplumber  # type: ignore

        records = []
        with pdfplumber.open(path) as pdf:
            for page_index, page in enumerate(pdf.pages):
                tables = page.extract_tables() or []
                for table_index, table in enumerate(tables):
                    if not table or len(table) < 2:
                        continue
                    frame = pd.DataFrame(table[1:], columns=table[0])
                    records.extend(frame_records(frame, f"pdf_page_{page_index + 1}_table_{table_index + 1}"))
        return records
    except Exception:
        pass

    try:
        from pypdf import PdfReader  # type: ignore

        reader = PdfReader(str(path))
        rows = []
        for page_index, page in enumerate(reader.pages):
            text = page.extract_text() or ""
            for line in text.splitlines():
                if line.strip():
                    rows.append({"text": line.strip(), "_source": f"pdf_page_{page_index + 1}"})
        return rows
    except Exception as error:
        raise RuntimeError("PDF parsing requires pdfplumber or pypdf. Add one of them to requirements.txt.") from error


def read_file(path):
    suffix = path.suffix.lower()

    if suffix in {".xlsx", ".xlsm", ".xltx", ".xltm", ".xls", ".ods"}:
        sheets = pd.read_excel(path, sheet_name=None)
        records = []
        for sheet_name, frame in sheets.items():
            records.extend(frame_records(frame, f"sheet:{sheet_name}"))
        return records

    if suffix in {".csv"}:
        return frame_records(pd.read_csv(path), "csv")

    if suffix in {".tsv", ".tab"}:
        return frame_records(pd.read_csv(path, sep="\t"), "tsv")

    if suffix in {".txt"}:
        return frame_records(pd.read_csv(path, sep=None, engine="python"), "txt")

    if suffix in {".json"}:
        return frame_records(pd.read_json(path), "json")

    if suffix in {".html", ".htm"}:
        records = []
        for index, frame in enumerate(pd.read_html(path)):
            records.extend(frame_records(frame, f"html_table_{index + 1}"))
        return records

    if suffix in {".parquet"}:
        return frame_records(pd.read_parquet(path), "parquet")

    if suffix in {".feather"}:
        return frame_records(pd.read_feather(path), "feather")

    if suffix in {".pkl", ".pickle"}:
        return frame_records(pd.read_pickle(path), "pickle")

    if suffix in {".pdf"}:
        return read_pdf(path)

    raise RuntimeError(f"Unsupported file type: {suffix or 'unknown'}")


def main():
    if len(sys.argv) < 2:
        raise RuntimeError("Missing file path.")

    path = Path(sys.argv[1])
    records = read_file(path)
    print(json.dumps({"rows": records[:15000], "rowCount": len(records)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
