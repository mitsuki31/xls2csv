import csv
import sys
from pathlib import Path
from typing import Optional, Set, List, Tuple

from openpyxl import load_workbook

from xls2csv.utils import DEFAULT_TEMPLATE, PathLike, format_output_name

# .xls is not supported due to openpyxl limitation and that's legacy format anyway
SUPPORTED_EXTS: Set[str] = { ".xlsx", ".xlsb", ".xlsm" }

def convert_single(
    excel_file: PathLike,
    output: Optional[PathLike] = None,
    /,
    *,
    sheet: Optional[str] = None,
    all_sheets: bool = False,
    template: Optional[str] = None,
    overwrite: bool = False
) -> None:
    """
    Convert a single Excel file to CSV format.

    Args:
        excel_file (str or Path-like object): Path to the Excel file to convert.
        output (str or Path-like object, optional): Path to the output CSV file.
        sheet (str, optional): Name of the sheet to convert. Default to active sheet.
        all_sheets (bool, optional): Whether to convert all sheets. Default to active sheet.
        template (str, optional): Template for the output file name. Default to `"%(name)-[%(sheet)].%(ext)"`.
        overwrite (bool, optional): Whether to overwrite existing files. Default to False.

    Raises:
        FileNotFoundError: If the Excel file is not found.
        ValueError: If the sheet name is not found in the Excel file.

    Returns:
        None
    
    Notes:
        - If `output` is a folder, the CSV file will be saved in that folder.
        - If `output` is a file, the CSV file will be saved in that file.
        - If `output` is not specified, the CSV file will be saved in the same folder as the Excel file.
        - If `all_sheets` is set to True, all sheets will be converted to CSV files.
        - If both `all_sheets` and `sheet` are specified, `all_sheets` will take precedence over sheet.
    """
    excel_file = Path(excel_file)

    if not excel_file.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_file}")

    wb = load_workbook(excel_file, data_only=True, read_only=True)
    try:
        if all_sheets:  # Take priority over sheet
            sheets = wb.sheetnames
        elif sheet:
            if sheet not in wb.sheetnames:
                raise ValueError(f"Sheet '{sheet}' not found in '{excel_file.name}'")
            sheets = [sheet]
        else:
            if not wb.sheetnames:
                raise ValueError(f"No sheets found in '{excel_file.name}'")
            sheets = [wb.active.title if wb.active else wb.sheetnames[0]]

        output_path = Path(output) if output is not None else None
        output_is_file = output_path and output_path.suffix.lower() == ".csv"

        if output_is_file and len(sheets) > 1:
            raise ValueError("A single output CSV path cannot be used with multiple sheets")

        template = template or DEFAULT_TEMPLATE

        if output_path and not output_is_file:
            output_path.mkdir(parents=True, exist_ok=True)

        for sheet_name in sheets:
            ws = wb[sheet_name]

            # --- filename generation ---
            filename = format_output_name(
                template,
                file=excel_file,
                sheet=sheet_name,
                ext="csv"
            )

            # --- resolve target path ---
            if output_path is None:
                target = excel_file.parent / filename
            elif output_is_file:
                target = output_path
            else:
                target = output_path / filename

            if target.exists() and not overwrite:
                raise FileExistsError(f"File '{target}' already exists. Use --overwrite to overwrite it.")

            with target.open("w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                for row in ws.iter_rows(values_only=True):
                    writer.writerow(row)
            print(f"Converted: '{excel_file}' ({sheet_name}) -> '{target}'")
    finally:
        wb.close()

def convert_batch(
    folder: PathLike,
    output: Optional[PathLike] = None,
    /,
    *,
    sheet: Optional[str] = None,
    all_sheets: bool = False,
    template: Optional[str] = None,
    overwrite: bool = False
) -> None:
    """
    Convert all Excel files in a folder to CSV format.

    Args:
        folder (str or Path-like object): Path to the folder containing Excel files.
        output (str or Path-like object, optional): Path to the output folder. Default to the same folder as the Excel files.
        sheet (str, optional): Name of the sheet to convert. Default to active sheet.
        all_sheets (bool, optional): Whether to convert all sheets. Default to active sheet.
        template (str, optional): Template for the output file name. Default to `"%(name)-[%(sheet)].%(ext)"`.
        overwrite (bool, optional): Whether to overwrite existing files. Default to False.

    Raises:
        FileNotFoundError: If the folder is not found.
        ValueError: If the sheet name is not found in the Excel file.

    Returns:
        None

    Notes:
        - If `output` is a folder, the CSV file will be saved in that folder.
        - If `output` is not specified, the CSV file will be saved in the same folder.
        - If `all_sheets` is set to True, all sheets will be converted to CSV files.
        - If both `all_sheets` and `sheet` are specified, `all_sheets` will take precedence over sheet.
    """
    folder = Path(folder)

    if not folder.exists():
        raise FileNotFoundError(f"Folder not found: {folder}")
    if not folder.is_dir():
        raise ValueError(f"Expected a directory, got: {folder}")

    output_path = Path(output) if output else folder

    if output_path.exists() and not output_path.is_dir():
        raise ValueError("Batch output must be a directory")

    output_path.mkdir(parents=True, exist_ok=True)

    excel_files = [
        f for f in folder.iterdir()
        if f.is_file() and f.suffix.lower() in SUPPORTED_EXTS
    ]

    if not excel_files:
        raise ValueError(f"No Excel files found in: {folder}")

    errors: List[Tuple[Path, Exception]] = []

    for excel_file in excel_files:
        try:
            print(f"Processing: {excel_file.name}")
            convert_single(
                excel_file,
                output_path,
                sheet=sheet,
                all_sheets=all_sheets,
                template=template,
                overwrite=overwrite
            )
        except Exception as e:
            errors.append((excel_file, e))

    if errors:
        message = "\n".join(
            f"- {file.name}: {err.__class__.__name__}: {err}"
            for file, err in errors
        )
        raise RuntimeError(f"Batch conversion failed:\n{message}")