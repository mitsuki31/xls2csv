from datetime import datetime
from pathlib import Path
from typing import (
    Callable,
    Generator,
    LiteralString,
    Optional,
    Set,
    Tuple,
    TypeAlias,
    Union,
)

PathLike: TypeAlias = Union[str, Path]

DEFAULT_TEMPLATE: LiteralString = "%(name)-[%(sheet)].%(ext)"
SUPPORTED_PLACEHOLDERS: Set[LiteralString] = {
    "%(name)", "%(ext)", "%(sheet)",
    "%(date)", "%(year)", "%(month)",
    "%(day)", "%(day_s)",
}

def sanitize_filename(name: str) -> str:
    """
    Convert a string to a safe filename.

    Args:
        name (str): The string to convert.

    Returns:
        str: The safe filename.
    """
    invalid = '<>:"/\\|?*'
    return "".join("_" if ch in invalid else ch for ch in name).strip()

def format_output_name(
    template: str,
    *,
    file: PathLike,
    ext: Optional[str] = None,
    sheet: Optional[str] = None,
) -> str:
    """
    Format output filename using a limited placeholder system.

    Supported placeholders:
        `%(name)`  -> file name without extension
        `%(ext)`   -> output file extension (without dot)
        `%(sheet)` -> sheet name (sanitized)
        `%(date)`  -> current date in `YYYY-MM-DD` format
        `%(year)`  -> current year in `YYYY` format
        `%(month)` -> current month in `MM` format
        `%(day)`   -> current day in `DD` format
        `%(day_s)`  -> current day of week in `DDD` format (e.g., Mon)

    Args:
        template (str): Template string.
        file (PathLike): Source Excel file.
        ext (str, optional): Output file extension (without dot). Default to extension of `file`.
        sheet (str, optional): Sheet name.

    Returns:
        str: Formatted filename (not full path).
    """
    file = Path(file)
    name = file.stem
    ext = ext or file.suffix.lstrip(".")
    # Sanitize sheet name because it will be used in filename
    sheet_safe = sanitize_filename(sheet) if sheet else ""
    today_dt = datetime.now()
    date_str = today_dt.strftime("%Y-%m-%d")  # YYYY-MM-DD
    year_str = today_dt.strftime("%Y")  # YYYY
    month_str = today_dt.strftime("%m")  # MM
    day_str = today_dt.strftime("%d")  # DD
    day_s_str = today_dt.strftime("%a")  # Mon, Tue, Wed, etc...

    result = template
    result = result                             \
        .replace("%(name)", name)               \
        .replace("%(ext)", ext)                 \
        .replace("%(date)", date_str)           \
        .replace("%(year)", year_str)           \
        .replace("%(month)", month_str)         \
        .replace("%(day)", day_str)             \
        .replace("%(day_s)", day_s_str)

    if "%(sheet)" in result:
        if sheet is None:
            raise ValueError("Template requires '%(sheet)' but no sheet name was provided")
        result = result.replace("%(sheet)", sheet_safe)

    return result

def print_summary(from_path: Path, to_path: Path, sheet_name: str, total_rows: int) -> None:
    """
    Print summary of the conversion.

    Args:
        from_path (Path): Path to the source Excel file.
        to_path (Path): Path to the output CSV file.
        sheet_name (str): Name of the sheet converted.
        total_rows (int): Total number of rows written.
    """
    target_size = to_path.stat().st_size
    print(f"Converted: \"{from_path}::{sheet_name}\" -> \"{to_path}\"")
    print(f"  >> {total_rows} rows written")
    print(f"  >> Total size: {target_size / 1024:.2f} KB ({target_size} bytes)")

def is_temp_excel(path: Path, additional_checks: Tuple[Callable[[str], bool]] = ()) -> bool:
    """
    Check if a file is a temporary/lock file created by Excel editors.

    These files should be excluded from processing as they represent
    unsaved changes or lock files.

    Patterns checked:
    - `"~$"` prefix: Microsoft Excel temp files (Windows)
    - `"._"` prefix: LibreOffice temp files (Linux)
    - `".~lock.#"` suffix: macOS lock files

    Args:
        path (Path): The file path to check.
        additional_checks (tuple[Callable[[str], bool]], optional):
            A tuple of callable to perform additional checks on the filename.

    Returns:
        bool: True if the filename matches known temporary Excel file patterns.
    """
    name = path.name
    checks = (
        lambda n: n.startswith("~$"),  # Microsoft Excel (Windows)
        lambda n: n.startswith("._"),  # LibreOffice (Linux)
        lambda n: n.startswith(".~lock.") and n.endswith("#"),  # macOS Metadata
        # Add more filters here when necessary
        *additional_checks,
    )

    return any(check(name) for check in checks)

def is_excel_file(path: Path, exts: Set[str]) -> bool:
    """
    Check if a file is an Excel file.

    This function uses ``is_temp_excel`` to check if the file is NOT a temporary/lock
    file created by Excel editors, and ``exts`` to check if the file has a
    valid Excel extension.

    Args:
        path (Path): The file path to check.
        exts (Set[str]): Set of lowercase file extensions to match (e.g., `{”.xlsx”, ”.xlsb”}`).

    Returns:
        bool: True if the file is an Excel file.
    """
    return path.is_file() and path.suffix.lower() in exts and not is_temp_excel(path)

def iter_excels(root: Path, exts: Set[str]) -> Generator[Path, None, None]:
    """
    Recursively iterate over Excel files in directory tree, excluding temp/lock files.

    Args:
        root (Path): Root directory to search.
        exts (Set[str]): Set of lowercase file extensions to match (e.g., `{".xlsx", ".xlsb"}`).

    Yields:
        Path: Valid Excel file paths (non-temp).
    """
    for p in root.rglob("*"):
        if is_excel_file(p, exts):
            yield p
