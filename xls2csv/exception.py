from pathlib import Path
from typing import Optional, Sequence, Tuple, TypeVar, Union

ERROR_CODES = {
    "E_UNKNOWN": "E_UNKNOWN",
    "E_IO": "E_IO",
    "E_FILEEXIST": "E_EXIST",
    "E_UNSUPPORTED": "E_UNSUPPORTED",
    "E_NOTEXCELFILE": "E_NOTEXCELFILE",
    "E_SHEETNOTFOUND": "E_SHEETNOTFOUND",
    "E_INVALIDDATA": "E_INVALIDDATA",
    "E_BATCHPROCESSING": "E_BATCHPROCESSING",
}

# NOTE: NOT STABLE! May change in the future.
EXIT_CODES = {
    "SUCCESS": 0,
    "E_UNKNOWN": 1,
    "E_IO": 2,
    "E_FILEEXIST": 3,
    "E_UNSUPPORTED": 4,
    "E_NOTEXCELFILE": 5,
    "E_SHEETNOTFOUND": 6,
    "E_INVALIDDATA": 7,
    "E_BATCHPROCESSING": 111,
}

_BPE = TypeVar("_BPE", bound="BatchProcessingError")
class XLS2CSVError(RuntimeError):
    """
    Base exception class for XLS2CSV errors.

    Provides some useful properties (e.g., :attr:`code`, :attr:`path`, :attr:`filename`)
    for better error context.
    """
    def __init__(
        self, /,
        message: str,
        *,
        filepath: Optional[Union[str, Path]] = None,
        err_code: ERROR_CODES = ERROR_CODES["E_UNKNOWN"]
    ) -> None:
        """
        Initialize the exception.

        Args:
            message (str): Error message.
            filepath (str or Path, optional): Optional file path for context.
            err_code (str, optional): Error code.
        """
        super().__init__(self._format_message(path=filepath, message=message))
        self._path = Path(filepath) if filepath is not None else None
        self._err_code = ERROR_CODES.get(err_code, ERROR_CODES["E_UNKNOWN"])

    def _format_message(self, path: Optional[Union[str, Path]], message: str) -> str:
        if path is not None:
            return f"{message} [file: {Path(path)}]"
        return message

    def __str__(self) -> str:
        return self.args[0]  # Do not use formatted message from parent class

    @property
    def code(self) -> ERROR_CODES:
        """
        Error code.

        Returns:
            ERROR_CODES: An error code from ``ERROR_CODES`` enum.
        """
        return self._err_code

    @property
    def path(self) -> Optional[Path]:
        """
        Path of the associated file, if provided.

        Returns:
            Path or None: File path object.
        """
        return self._path

    @property
    def filename(self) -> Optional[str]:
        """
        File name that associated with, if provided.

        Returns:
            str or None: A string representing file name.
        """
        return self._path.name if self._path is not None else None

class NotAnExcelFileError(XLS2CSVError):
    """
    Raised when a file is not a recognized Excel format.
    """
    def __init__(self, message: Optional[str] = None, *, filepath: Optional[Union[str, Path]] = None) -> None:
        """
        Initialize the exception.

        Args:
            message (str, optional): Error message to use (overrides default).
            filepath (str or Path, optional): Optional file path for context.
        """
        super().__init__(
            message or "File does not seem to be an Excel file",
            filepath=filepath,
            err_code="E_NOTEXCELFILE"
        )

class UnsupportedFormatError(XLS2CSVError):
    """
    Raised when an Excel file format is not supported by the underlying reading engine.
    """
    def __init__(
        self, /,
        expected: Sequence[str],
        actual: Optional[str] = None,
        *,
        message: Optional[str] = None,
        filepath: Optional[Union[str, Path]] = None
    ) -> None:
        """
        Initialize the exception.

        Args:
            expected (Sequence[str]): Sequence of expected Excel file formats.
            actual (str, optional): Actual Excel file format, if not provided, it will be inferred from ``filepath``.
            message (str, optional): Error message to use (overrides default).
            filepath (str or Path, optional): Optional file path for context.
        """
        msg = f"{message or 'Excel file format not supported.'} Expected: {', '.join(expected)}"
        expected = [str(item).lower() for item in expected]

        if actual is not None:
            msg += f" (actual: {actual})"
        elif actual is None and filepath is not None:
            msg += f" (actual: {Path(filepath).suffix})"

        super().__init__(msg, filepath=filepath, err_code="E_UNSUPPORTED")
        self._expected_formats = tuple(set(expected))

    @property
    def expected_formats(self) -> Tuple[str, ...]:
        """
        Expected Excel file format.

        Returns:
            tuple[str]: Tuple of expected Excel file formats.
        """
        return self._expected_formats

class NoSheetFoundError(XLS2CSVError):
    """
    Raised when a specified sheet name is not found inside Excel file.
    """
    def __init__(
        self, /,
        sheet_name: str, 
        message: Optional[str] = None,
        *,
        filepath: Optional[Union[str, Path]] = None
    ) -> None:
        """
        Initialize the exception.

        Args:
            sheet_name (str): Name of the sheet that was not found.
            message (str, optional): Error message to use (overrides default).
            filepath (str or Path, optional): Optional file path for context.
        """
        super().__init__(
            message or f"Sheet '{sheet_name}' not found in the Excel file",
            filepath=filepath,
            err_code="E_SHEETNOTFOUND"
        )
        self._sheet_name = sheet_name

    @property
    def sheet_name(self) -> str:
        """
        Sheet name.

        Returns:
            str: Sheet name.
        """
        return self._sheet_name

class OutputFileExistsError(XLS2CSVError, FileExistsError):
    """
    Raised when an output file already exists, and `--overwrite` option is not enabled.
    """
    def __init__(self, message: Optional[str] = None, *, filepath: Optional[Union[str, Path]] = None) -> None:
        """
        Initialize the exception.

        Args:
            message (str, optional): Error message to use (overrides default).
            filepath (str or Path, optional): Optional file path for context.
        """
        super().__init__(
            message or "Output file already exists. Use --overwrite to overwrite it",
            filepath=filepath,
            err_code="E_FILEEXIST"
        )

class BatchProcessingError(XLS2CSVError):
    """
    Raised when an error occurs during batch processing.
    """
    def __init__(
        self, /,
        message: str,
        *,
        filepath: Optional[Union[str, Path]] = None,
        errors: Optional[Sequence[RuntimeError]] = None
    ) -> None:
        """
        Initialize the exception.

        Args:
            message (str): Error message.
            filepath (str or Path, optional): Optional file path for context.
            errors (Sequence[RuntimeError], optional): Sequence of errors that occurred during batch processing.
        """
        super().__init__(message, filepath=filepath, err_code="E_BATCHPROCESSING")
        self._errors = list(errors) if errors is not None else []

    def add_error(self: _BPE, err: RuntimeError) -> _BPE:
        """
        Add an error to the batch processing error.

        For example:

            err = BatchProcessingError("Batch processing failed")
            err.add_error(NotAnExcelFileError(filepath="/path/to/file"))

        Args:
            err (RuntimeError): Error to add.

        Returns:
            BatchProcessingError: self (for chaining).
        """
        self._errors.append(err)
        return self

    def append(self: _BPE, err: RuntimeError) -> _BPE:
        """
        Alias for ``add_error``.

        Add an error to the batch processing error.

        Args:
            err (RuntimeError): Error to add.

        Returns:
            BatchProcessingError: self (for chaining).
        """
        return self.add_error(err)

    @property
    def errors(self) -> Tuple[RuntimeError, ...]:
        """
        All errors that occurred during batch processing.

        Returns:
            tuple[RuntimeError]: Tuple of errors.
        """
        return tuple(self._errors)



def format_err(err: XLS2CSVError, /, *, with_code: bool = True, with_type: bool = False) -> str:
    """
    Helper function to format error from an ``XLS2CSVError`` instance.

    The file path is always included in the error message, unless ``err.path`` is ``None``.

    Example of formatted message:

        "[<ERROR_CODE>] <MESSAGE> [file: <PATH>]"

    or (when ``with_type`` is True):

        "[<ERROR_CODE>::<CLASS_NAME>] <MESSAGE> [file: <PATH>]"

    Args:
        err (XLS2CSVError): Error object, must be an instance of ``XLS2CSVError`` or its subclasses.
        with_code (bool, optional): Whether to include error code in the message. Defaults to True.
        with_type (bool, optional): Whether to include error class name in the message. Defaults to False.

    Returns:
        str: Formatted error message.
    """
    parts = []

    # error code (primary identifier)
    if with_code and getattr(err, "code", None):
        parts.append(err.code)

    # optional class name (debugging)
    if with_type:
        parts.append(type(err).__name__)

    prefix = f"[{'::'.join(parts)}] " if parts else ""
    message = f"{prefix}{str(err)}"

    return message
