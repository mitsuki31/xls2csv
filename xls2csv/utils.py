from pathlib import Path
from typing import Optional, TypeAlias, Union

PathLike: TypeAlias = Union[str, Path]

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
    sheet: Optional[str] = None,
) -> str:
    """
    Format output filename using a limited placeholder system.

    Supported placeholders:
        `%name%`  -> file name without extension
        `%ext%`   -> file extension (without dot)
        `%sheet%` -> sheet name (sanitized)

    Args:
        template (str): Template string.
        file (PathLike): Source Excel file.
        sheet (str, optional): Sheet name.

    Returns:
        str: Formatted filename (not full path).
    """
    file = Path(file)
    name = file.stem
    ext = file.suffix.lstrip(".")
    # Sanitize sheet name because will be used in filename
    sheet_safe = sanitize_filename(sheet) if sheet else ""

    result = template
    result = result.replace("%name%", name)
    result = result.replace("%ext%", ext)

    if "%sheet%" in result:
        if sheet is None:
            raise ValueError("Template requires '%sheet%' but no sheet was provided")
        result = result.replace("%sheet%", sheet_safe)

    return result