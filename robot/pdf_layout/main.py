from __future__ import annotations

from pathlib import Path

from .excel_writer import write_rows_to_excel
from .field_extractors import ExtractionResult, extract_fields_from_layout
from .pdf_reader import read_pdf_layout


def extract_pdf_result(pdf_path: str | Path) -> ExtractionResult:
    layout = read_pdf_layout(pdf_path)
    return extract_fields_from_layout(layout)


def extract_pdf_fields(pdf_path: str | Path) -> dict:
    result = extract_pdf_result(pdf_path)
    return result.values_dict()


def process_pdf_folder(pdf_dir: str | Path, excel_path: str | Path) -> Path:
    pdf_dir = Path(pdf_dir)
    rows: list[dict] = []
    for pdf_path in sorted(pdf_dir.glob("*.pdf")):
        row = {"archivoPdf": pdf_path.name}
        row.update(extract_pdf_fields(pdf_path))
        rows.append(row)
    return write_rows_to_excel(rows, excel_path)


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Procesa una carpeta de PDFs y genera un Excel final.")
    parser.add_argument("pdf_dir", help="Carpeta con PDFs")
    parser.add_argument("excel_path", help="Ruta del Excel de salida")
    args = parser.parse_args()
    output = process_pdf_folder(args.pdf_dir, args.excel_path)
    print(output)
