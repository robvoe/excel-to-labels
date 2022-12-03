from pathlib import Path
import argparse

import pandas as pd
from fpdf import FPDF


def excel_to_labels(input_xlsx: Path, output_pdf: Path, cell_width_mm: float = 40, cell_height_mm: float = 22,
                    cell_margin_mm: float = 3, cell_frame: bool = True) -> None:
    assert input_xlsx.is_file(), f"Input XLSX '{input_xlsx}' does not exist or is no file!"
    df: pd.DataFrame = pd.read_excel(input_xlsx, index_col=None, engine="openpyxl", convert_float=False, dtype=str)
    df.columns = [c.lower() for c in df.columns]

    # Drop those entries, that are all-nans or which shall not be printed (see "print?" column in xlsx)
    df = df.dropna(axis="index", how="all")  # --> drops all-nan rows
    if "comment" in df.columns:
        df = df.drop(columns="comment")
    if "print?" in df.columns:
        df = df.drop(df[df["print?"].isna()].index)
        df = df.drop(columns="print?")

    print(f"Parsed Excel file. Generating {len(df)} labels...")

    # Print our labels
    pdf_generator = FPDF(orientation="P", unit="mm", format="A4")
    pdf_generator.set_margins(left=20, top=10, right=10)
    pdf_generator.set_font("Arial", "B", 7)
    pdf_generator.add_page()
    x_start, y_start = pdf_generator.get_x(), pdf_generator.get_y()

    n_rows_per_page, n_cols_per_page = (pdf_generator.h-10) // cell_height_mm, (pdf_generator.w-20-10) // cell_width_mm
    n_labels_per_page = int(n_rows_per_page * n_cols_per_page)

    n_lines_per_cell = len(df.columns)
    line_height = (cell_height_mm - cell_margin_mm * 2) / (n_lines_per_cell + 1)

    for n_label, (_, row) in enumerate(df.iterrows()):
        if n_label > 0 and n_label % n_labels_per_page == 0:
            pdf_generator.add_page()
        if n_label > 0 and n_label % n_cols_per_page == 0:
            pdf_generator.ln()
        n_row, n_col = n_label % n_labels_per_page // n_cols_per_page, n_label % n_labels_per_page % n_cols_per_page
        x_cell, y_cell = x_start + n_col * cell_width_mm, y_start + n_row * cell_height_mm
        pdf_generator.set_xy(x=x_cell, y=y_cell)
        pdf_generator.cell(cell_width_mm, cell_height_mm, txt="", border=cell_frame)
        for i_line, (_, txt) in enumerate(row.iteritems()):
            pdf_generator.text(x=x_cell + cell_margin_mm, y=y_cell + cell_margin_mm + (i_line+1)*line_height, txt=txt)

    pdf_generator.output(str(output_pdf))
    print(f"Done -- '{output_pdf}' contains {(len(df)//n_labels_per_page)+1} page(s).")


if __name__ == '__main__':
    _default_input_file = Path(__file__).parent / "template.xlsx"
    try:
        _default_input_file = _default_input_file.relative_to(Path.cwd())
    except ValueError:
        pass

    _parser = argparse.ArgumentParser(description="Converts an Excel table into a batch of labels, as a PDF file.")
    _parser.add_argument("input", default=str(_default_input_file), type=str, nargs="?",
                         help="The input Excel file. Should follow the format of 'template.xlsx'.")
    _parser.add_argument("output", type=str, default="output.pdf", nargs="?", help="The output PDF file.")
    _parser.add_argument("-width", "--cell_width_mm", type=float, default=40,
                         help="Width (mm) of each single label cell.")
    _parser.add_argument("-height", "--cell_height_mm", type=float, default=22,
                         help="Height (mm) of each single label cell.")
    _parser.add_argument("-margin", "--cell_margin_mm", type=float, default=3,
                         help="Margin (mm) between text and label cell frame.")
    _parser.add_argument("--no_frame", action="store_true", help="No frame around label cells.")
    args = _parser.parse_args()

    excel_to_labels(input_xlsx=Path(args.input), output_pdf=Path(args.output), cell_width_mm=args.cell_width_mm,
                    cell_height_mm=args.cell_height_mm, cell_margin_mm=args.cell_margin_mm, cell_frame=not args.no_frame)
