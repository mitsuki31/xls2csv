#!/usr/bin/env python3

from __future__ import annotations

import click
from pathlib import Path
from typing import Optional

from xls2csv.converter import (
    convert_single,
    convert_batch,
    DEFAULT_TEMPLATE,
    SUPPORTED_EXTS
)

__VERSION__ = "0.1.0"

@click.command(context_settings={"help_option_names": ["-h", "--help"]})
@click.argument("input_path", type=click.Path(path_type=Path))
@click.option(
    "-o", "--output",
    type=click.Path(path_type=Path),
    help="Output file or directory"
)
@click.option(
    "-t", "--template",
    type=str,
    default=DEFAULT_TEMPLATE,
    help=f"Template for the output file name. Default to \"{DEFAULT_TEMPLATE}\""
)
@click.option("-s", "--sheet", help="Specific sheet name")
@click.option("--all-sheets", is_flag=True, help="Export all sheets")
@click.option("-f", "--force", "--overwrite", is_flag=True, help="Overwrite existing files")
@click.version_option(
    __VERSION__,
    "-v",
    "--version",
    prog_name="xls2csv",
    message="-- %(prog)s v%(version)s --"
)
def cli(
    input_path: Path,
    output: Optional[Path],
    template: Optional[str],
    sheet: Optional[str],
    all_sheets: bool,
    force: bool
):
    """Convert Excel file(s) to CSV.

    INPUT_PATH can be a file (single conversion) or a folder (batch conversion).
    """

    # --- validation ---
    if not input_path.exists():
        raise click.ClickException(f"Path not found: {input_path}")

    if sheet and all_sheets:
        raise click.BadArgumentUsage("Use either --sheet or --all-sheets, not both")

    # --- dispatch ---
    if input_path.is_dir():
        # Batch mode
        if output and output.suffix:
            raise click.BadOptionUsage("--output", "Batch output must be a directory")

        try:
            convert_batch(
                input_path,
                output,
                template=template,
                sheet=sheet,
                all_sheets=all_sheets,
                overwrite=force
            )
        except Exception as e:
            raise click.ClickException(f"[{e.__class__.__name__}] {e}")
    else:
        # Single file mode
        if not input_path.is_file():
            raise click.FileError(str(input_path), "File cannot be processed")

        # Raise error to convert first if the file is .xls (legacy format)
        if input_path.suffix.lower() == ".xls":
            raise click.ClickException(
                "Invalid format: .xls is a legacy and is not supported by this tool. Please convert it to .xlsx first."
            )

        if input_path.suffix.lower() not in SUPPORTED_EXTS:
            raise click.ClickException(
                f"Unsupported file type: {input_path.suffix}\n"
                f"Supported types: {', '.join(SUPPORTED_EXTS)}"
            )

        try:
            convert_single(
                input_path,
                output,
                template=template,
                sheet=sheet,
                all_sheets=all_sheets,
                overwrite=force
            )
        except Exception as e:
            raise click.ClickException(f"[{e.__class__.__name__}] {e}")

if __name__ == "__main__":
    cli()