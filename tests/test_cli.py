import csv
from pathlib import Path

import pytest
from click.testing import CliRunner

from xls2csv.cli import __VERSION__, cli

DATA_DIR = Path(__file__).parent / "data"


def run_cli(runner, args):
    result = runner.invoke(cli, args)
    assert result.exit_code == 0, result.output
    return result

def get_single_csv(tmp_path: Path) -> Path:
    files = list(tmp_path.glob("*.csv"))
    assert len(files) == 1
    return Path(files[0])

def read_csv(path: Path):
    with path.open(newline="", encoding="utf-8") as f:
        return list(csv.reader(f))

@pytest.fixture
def runner():
    return CliRunner()


# --- CLI VERSION TEST ---

def test_cli_version(runner):
    result = run_cli(runner, ["--version"])
    assert result.exit_code == 0
    assert __VERSION__ in result.output.strip()


# --- BASIC TESTS ---

def test_basic_file_conversion(runner, tmp_path):
    input_file = DATA_DIR / "basic-data.xlsx"

    run_cli(runner, [str(input_file), "-o", str(tmp_path)])

    result_csv = get_single_csv(tmp_path)

    assert result_csv.stat().st_size > 0

    rows = read_csv(result_csv)
    assert rows[0] == ["id", "name", "amount", "date"]
    assert len(rows) > 1


def test_multi_sheet_conversion(runner, tmp_path):
    input_file = DATA_DIR / "basic-multi-sheets.xlsx"

    run_cli(runner, [str(input_file), "-o", str(tmp_path), "--all-sheets"])

    files = list(tmp_path.glob("*.csv"))
    assert len(files) >= 2  # depends on fixture


# --- EDGE CASES ---

def test_empty_file(runner, tmp_path):
    input_file = DATA_DIR / "empty-file.xlsx"

    result = runner.invoke(cli, [str(input_file), "-o", str(tmp_path)])

    result_csv = get_single_csv(tmp_path)
    assert result.exit_code == 0  # Still success, but ...
    assert result_csv.stat().st_size == 0  # the file is empty


def test_header_only_file(runner, tmp_path):
    input_file = DATA_DIR / "header-only.xlsx"

    run_cli(runner, [str(input_file), "-o", str(tmp_path)])

    result_csv = get_single_csv(tmp_path)
    assert result_csv.stat().st_size > 0

    rows = read_csv(result_csv)
    assert len(rows) == 1  # only header


def test_empty_null_data(runner, tmp_path):
    input_file = DATA_DIR / "empty-null-data.xlsx"

    run_cli(runner, [str(input_file), "-o", str(tmp_path)])

    result_csv = get_single_csv(tmp_path)
    assert result_csv.stat().st_size > 0

    rows = read_csv(result_csv)
    # N/A will be replaced with NULL in CSV by default
    assert any("" in row or "NULL" in row for row in rows)


# --- DATA VALIDATION ---

def test_special_characters(runner, tmp_path):
    input_file = DATA_DIR / "special-chars.xlsx"

    run_cli(runner, [str(input_file), "-o", str(tmp_path)])

    result_csv = get_single_csv(tmp_path)
    assert result_csv.stat().st_size > 0

    content = result_csv.read_text(encoding="utf-8")

    assert "André" in content
    assert "李雷" in content
    assert "أحمد" in content


def test_large_numbers(runner, tmp_path):
    input_file = DATA_DIR / "large-numbers.xlsx"

    run_cli(runner, [str(input_file), "-o", str(tmp_path)])

    result_csv = get_single_csv(tmp_path)
    assert result_csv.stat().st_size > 0

    content = result_csv.read_text("utf-8")

    assert "999999999999999" in content


def test_long_text(runner, tmp_path):
    input_file = DATA_DIR / "long-text.xlsx"

    run_cli(runner, [str(input_file), "-o", str(tmp_path)])

    result_csv = get_single_csv(tmp_path)
    assert result_csv.stat().st_size > 0

    content = result_csv.read_text("utf-8")

    assert "very long text field" in content.lower()


def test_multiple_data_types(runner, tmp_path):
    input_file = DATA_DIR / "multiple-data-types.xlsx"

    run_cli(runner, [str(input_file), "-o", str(tmp_path)])

    result_csv = get_single_csv(tmp_path)
    assert result_csv.stat().st_size > 0

    rows = read_csv(result_csv)

    header = rows[0]
    assert "active" in header
    assert "balance" in header


# --- DOMAIN TEST ---

def test_financial_data(runner, tmp_path):
    input_file = DATA_DIR / "financial-data.xlsx"

    run_cli(runner, [str(input_file), "-o", str(tmp_path)])

    result_csv = get_single_csv(tmp_path)
    assert result_csv.stat().st_size > 0

    rows = read_csv(result_csv)

    header = rows[0]
    assert "debit" in header
    assert "credit" in header
    assert "balance" in header
