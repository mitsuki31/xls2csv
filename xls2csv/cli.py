#!/usr/bin/env python3

import click
from pathlib import Path
from typing import Optional

from xls2csv.converter import (
    convert_single,
    convert_batch,
    SUPPORTED_EXTS
)

@click.command(context_settings={"help_option_names": ["-h", "--help"]})
@click.argument("input_path", type=click.Path(path_type=Path))
@click.option(
    "-o", "--output",
    type=click.Path(path_type=Path),
    help="Output file or directory"
)
@click.option("-s", "--sheet", help="Specific sheet name")
@click.option("--all-sheets", is_flag=True, help="Export all sheets")
def cli(
    input_path: Path,
    output: Optional[Path],
    sheet: Optional[str],
    all_sheets: bool
):
    """Convert Excel file(s) to CSV.

    INPUT_PATH can be a file (single conversion) or a folder (batch conversion).
    """

    # --- validation ---
    if not input_path.exists():
        raise click.BadParameter(f"Path not found: {input_path}")

    if sheet and all_sheets:
        raise click.BadParameter("Use either --sheet or --all-sheets, not both")

    # --- dispatch ---
    if input_path.is_dir():
        # Batch mode
        if output and output.suffix:
            raise click.BadParameter("Batch output must be a directory")

        convert_batch(
            input_path,
            output,
            sheet=sheet,
            all_sheets=all_sheets
        )

    else:
        # Single file mode
        if not input_path.is_file():
            raise click.BadParameter(f"Invalid file: {input_path}")

        # Raise error to convert first if the file is .xls
        if input_path.suffix.lower() == ".xls":
            raise click.BadParameter(
                "Format .xls is a legacy and is not supported by this tool. Please convert it to .xlsx first."
            )

        convert_single(
            input_path,
            output,
            sheet=sheet,
            all_sheets=all_sheets
        )

if __name__ == "__main__":
    cli()