# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

xlsx2csv is a Python utility that converts XLSX files to CSV format. It's designed to handle large XLSX files efficiently and supports multiple Python versions (2.4 to 3.14). The entire converter is implemented in a single Python file `xlsx2csv.py` (~54KB).

## Development Commands

### Testing
```bash
python3 test/run
```
This runs the comprehensive test suite that compares output from various test XLSX files against expected CSV files. Tests cover edge cases like datetime formatting, empty rows, hyperlinks, various delimiters, and multi-sheet files.

### Building
```bash
python3 -m build
```
Uses the modern Python build system defined in `pyproject.toml`.

### Installation for Development
```bash
pip install -e .
```
Install in editable mode for local development.

## Code Architecture

### Core Classes
- **Xlsx2csv**: Main converter class that orchestrates the conversion process
- **Workbook**: Represents the XLSX workbook structure and handles ZIP file extraction
- **Sheet**: Handles individual worksheet conversion with SAX parsing for memory efficiency
- **SharedStrings**: Manages the shared strings table
- **Styles**: Handles cell formatting and number formats
- **ContentTypes**: Manages XLSX content type definitions
- **Relationships**: Handles XLSX relationship mappings

### Key Design Patterns
- **SAX Parsing**: Uses `xml.parsers.expat` for memory-efficient XML parsing of large files
- **Streaming Processing**: Processes XLSX files without loading entire content into memory
- **Format Detection**: Comprehensive format mapping system (`FORMATS` and `STANDARD_FORMATS` dicts) for proper type conversion
- **Command-line Interface**: Uses argparse for Python 3+ and optparse fallback for Python 2.4

### File Structure
- `xlsx2csv.py`: Single-file implementation containing all classes and logic
- `test/`: Contains test XLSX/CSV file pairs and the test runner
- `test/run`: Python script that compares converter output against expected results

## Testing Strategy
The test suite uses a comparison-based approach:
1. Converts test XLSX files using the converter
2. Compares output with pre-generated expected CSV files
3. Tests both file input and STDIN input modes
4. Covers various edge cases: datetime formatting, hyperlinks, empty cells, multi-sheet files, different encodings

## Key Features Supported
- Multiple output formats and delimiters
- Date/time formatting with custom patterns
- Hyperlink extraction
- Multi-sheet processing
- Large file handling via streaming
- Cross-platform compatibility (Windows/Linux/macOS)
- Python 2.4 to 3.14 compatibility
