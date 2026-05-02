# xls2csv

A minimal CLI tool to convert Excel files into CSV format written in Python v3.

Designed for reliability in operational environments where redundant, portable backups (CSV) are required alongside primary spreadsheet workflows.

---

## Overview

In many office and accounting workflows, data is often stored in Excel and synced via cloud providers (e.g. OneDrive). While convenient, this setup introduces a single point of failure.

This tool provides a simple fallback mechanism:

- Convert Excel files into plain CSV
- Export all sheets or specific ones
- Batch process entire directories
- Generate consistent, scriptable outputs

The focus is not feature richness, but **predictable, safe conversion**.

> [!IMPORTANT]  
> In case you didn't know, converting Excel files to CSV format **will lose the formatting, formulas, etc**.
>
> It will only stores the calculated values of the cells.

By converting to CSV, you can have an emergency backup of your Excel data that can be saved locally (e.g., USB drive).

---

## Features

- **Lightweight** – Designed with minimal dependencies for easy setup and maintenance
- **Excel to CSV Conversion** – Convert a single Excel file into CSV format
- **Batch Processing** – Convert entire folders of Excel files in one command

### Supported Excel Files

| File Type                     | Extension | Supported |
| ----------------------------- | :-------- | :-------: |
| Excel 97-2003 Workbook        | `.xls`    | ❌        |
| Excel Workbook                | `.xlsx`   | ✅        |
| Excel Macro-Enabled Workbook  | `.xlsm`   | ✅        |
| Excel Binary Workbook         | `.xlsb`   | ✅        |
| Excel Open XML Spreadsheet    | `.ods`    | ❌        |

### Export Options

- Active sheet (default behavior)
- Specific sheet using `--sheet`
- All sheets using `--all-sheets`

### Output Flexibility

- Save to a file in single-file mode
- Save to a directory in batch mode
- Optional template-based naming for generated files (`--template`)

### Additional Benefits

- **Cross-platform support** – Works seamlessly on Windows, Linux, and macOS
- **Memory efficient** – Uses read-only processing to keep memory usage low

---

## Project Inspiration

This project originated from a practical issue encountered while managing financial data backups in a local workflow.

> [!IMPORTANT]  
> Please take as important note, **DO NOT** upload any sensitive files such as company financial data
> either to public or private repository which may be violating the company policy and I'm not responsible for that.

An initial approach involved using Git for versioning Excel files. However, since Excel formats (e.g., `.xlsx`) are binary, even small changes resulted in large diffs and rapid growth of the `.git` directory when committed directly without [**Git LFS**](https://git-lfs.com/). This made the repository inefficient and difficult to maintain over time.

To address this, a simple conversion step was introduced: transform Excel files into CSV before committing. Unlike binary formats, CSV files are text-based, enabling meaningful diffs, smaller repository size, and better traceability of changes.

What began as a small utility script evolved into a dedicated CLI tool focused on one purpose—providing a reliable way to generate lightweight, version-friendly backups from spreadsheet data.

---

## Installation

### Prerequisites

Ensure you have [**Python 3.10**](https://www.python.org/downloads/) ([see release notes](https://www.python.org/downloads/release/python-31020/)) or higher installed.

That's all you need, really.

### Option 1 — Recommended (`pipx`)

[`pipx`](https://pipx.pypa.io/stable/) is a tool for installing and running Python applications in isolated environments. Best for Ubuntu 23.04+, Debian 12+, and Fedora 38+ (these are distros that adopts [PEP 668](https://peps.python.org/pep-0668/)).

#### Install `pipx`

For Unix users:

```bash
sudo apt install pipx
```

> [!NOTE]  
> For **Termux** (Android) users, you don't need to use `sudo` to install package.  
> Just run `pkg install pipx` instead.

For Windows users:

```powershell
python -m pip install pipx && python -m pipx ensurepath
```

> Or, check [`pipx` installation guide](https://pipx.pypa.io/stable/how-to/install-pipx/) for more information on how to install `pipx`.

#### Install `xls2csv`

```bash
pipx install git+https://github.com/mitsuki31/xls2csv.git@latest
```

#### Verify installation

```bash
xls2csv --version
```

---

### Option 2 — `pip` (global or user install)

> [!WARNING]  
> Recommended to use `pipx` instead, `pip` can sometimes cause dependency conflicts in some environments.

```bash
pip install git+https://github.com/mitsuki31/xls2csv.git@latest
```
> You may need `--user` depending on your environment.

After installation:

```bash
xls2csv --version
```

---

## Usage

### Basic syntax

```bash
xls2csv [OPTIONS] INPUT_PATH
```

`INPUT_PATH` determines behavior:

- File → single conversion
- Directory → batch conversion

See [options](#options--behavior) for more detailed information.

> [!NOTE]  
> - Sheet names are sanitized for filesystem compatibility
>   - For example: `April/2026` → `April_2026`
> - Empty or invalid Excel files will raise errors
> - CSV encoding is set to UTF-8
> - Read-only Excel loading for lower memory usage

---

## Examples

### Convert single file

```bash
xls2csv report.xlsx
```

> [!NOTE]  
> If you don't specify the output file, it will be saved in the same folder as the input file.

### Convert to specific output file

```bash
xls2csv report.xlsx -o output.csv
```

> [!NOTE]  
> The `-o` / `--output` option accepts either a file path or a directory:  
> - If a file extension is provided (e.g. `a/b/c.csv`), the output is written to that file.  
> - If no extension is provided, the value is treated as a directory.

### Convert all sheets

```bash
xls2csv report.xlsx --all-sheets
```

### Convert specific sheet

```bash
xls2csv report.xlsx -s 2026-01
```

### Batch convert a folder

```bash
xls2csv ./reports
```

### Batch with output directory

```bash
xls2csv ./reports -o ./out
```

### Convert with template

```bash
xls2csv report.xlsx -t 'output-%(name)-%(sheet).csv'
```

---

## Options & Behavior

| Option                         | Description                                                            |
| ------------------------------ | ---------------------------------------------------------------------- |
| `-o, --output`                 | Output file (single mode) or directory                                 |
| `-t, --template`               | Template for output file name (default: `'%(name)-[%(sheet)].%(ext)'`) |
| `-s, --sheet`                  | Convert specific sheet                                                 |
| `--all-sheets`                 | Export all sheets                                                      |
| `-f`, `--force`, `--overwrite` | Overwrite existing files                                               |
| `-v`, `--version`              | Show version information                                               |
| `--help`                       | Show help message                                                      |

---

### Single file mode

| Input         | Output                |
| ------------- | --------------------- |
| No `-o`       | Same folder as source |
| `-o file.csv` | Exact file            |
| `-o folder/`  | File inside folder    |

---

### Batch mode

| Input         | Output                |
| ------------- | --------------------- |
| No `-o`       | Same folder           |
| `-o folder/`  | All files saved there |
| `-o file.csv` | ❌ Invalid             |

---

### Output File Placeholder

You can specify a placeholder in the output file name using `--template` option.

| Placeholder  | Description                                              |
|--------------|----------------------------------------------------------|
| `%(name)`    | The filename of the input file                          |
| `%(sheet)`   | The sheet name of the input file                        |
| `%(ext)`     | The extension of the output file (always `csv`)         |
| `%(date)`    | The current date in `YYYY-MM-DD` format                 |
| `%(year)`    | The current year in `YYYY` format                       |
| `%(month)`   | The current month in `MM` format                        |
| `%(day)`     | The current day in `DD` format                          |
| `%(day_s)`   | The current day of the week in `DDD` format (e.g., Mon) |

Example:

```bash
xls2csv report.xlsx -t 'output-%(name)-%(sheet).%(ext)'
```

> [!TIP]  
> The `%(date)` placeholder is a shortcut for `'%(year)-%(month)-%(day)'`.

> [!NOTE]  
> Default template string for the output file is `'%(name)-[%(sheet)].%(ext)'`.
>
> Example output file:
> ```text
> report-[Sheet1].csv
> ```

---

## Development

### Setup virtual environment

- Unix
  ```bash
  ./setup.sh && . .venv/bin/activate
  ```

- Windows (PowerShell)
  ```pwsh
  .\setup.ps1
  ```

### Install dependencies

```bash
pip install -e '.[dev]'
```

### Run the CLI

```bash
xls2csv --version
```

---

## Testing

All tests are located in the `tests/` directory. To run all tests, use `pytest`:

```bash
pytest -vv
```

---

## License

&copy; 2026 [Ryuu Mitsuki](https://github.com/mitsuki31).
Licensed under the [MIT License](./LICENSE).
