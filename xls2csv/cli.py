#!/usr/bin/env python3

from __future__ import annotations

import sys
from pathlib import Path
from typing import Optional

import click

from xls2csv.converter import (
    convert_single,
    convert_batch,
    DEFAULT_TEMPLATE,
    SUPPORTED_EXTS
)
from xls2csv.exception import (
    ERROR_CODES,
    format_err,
    BatchProcessingError,
    NotAnExcelFileError,
    UnsupportedFormatError,
    XLS2CSVError,
)

__VERSION__ = "0.2.0"

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

    This tool can automatically detect what mode to run based on the input path.
    If the input path is a directory, it will run in batch mode.
    Otherwise, it will run in single mode.
    """

    # --- validation ---
    if not input_path.exists():
        raise XLS2CSVError(
            f"No such file or directory: {input_path.resolve()}",
            err_code=ERROR_CODES["E_IO"]
        )

    if sheet and all_sheets:
        raise click.BadOptionUsage("--sheet", "Use either --sheet or --all-sheets, not both")

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
        except BatchProcessingError as be:
            errors = be.errors
            errors_msg = f"Batch conversion failed: {len(errors)} error(s) occurred\n"
            for error in errors:
                errors_msg += f"- {format_err(error, with_type=False)}\n"
            raise click.ClickException(errors_msg)
    else:
        # Single file mode
        if not input_path.is_file():
            raise NotAnExcelFileError("Unable to process the given file", filepath=input_path)

        # Raise error to convert first if the file is .xls (legacy format)
        if input_path.suffix.lower() == ".xls":
            raise UnsupportedFormatError(
                expected=SUPPORTED_EXTS,
                actual=input_path.suffix,
                message="Invalid format: .xls is a legacy and is not supported by this tool. Please convert it to .xlsx first.",
                filepath=input_path
            )

        if input_path.suffix.lower() not in SUPPORTED_EXTS:
            raise UnsupportedFormatError(
                expected=SUPPORTED_EXTS,
                actual=input_path.suffix,
                filepath=input_path
            )

        convert_single(
            input_path,
            output,
            template=template,
            sheet=sheet,
            all_sheets=all_sheets,
            overwrite=force
        )

if __name__ == "__main__":
    try:
        # pylint: disable=no-value-for-parameter
        cli()
    # pylint: disable=broad-except
    except Exception as e:
        print(format_err(e, with_code=True), file=sys.stderr)
        sys.exit(1)
