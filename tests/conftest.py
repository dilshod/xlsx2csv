"""Shared pytest fixtures for xlsx2csv tests.

This module provides common fixtures that can be used across all test files.
"""
import os
import tempfile
from pathlib import Path
from typing import Generator

import pytest


@pytest.fixture
def temp_dir() -> Generator[Path, None, None]:
    """Create a temporary directory for test files.

    Yields:
        Path: Path to the temporary directory.

    The directory and all its contents are automatically cleaned up after the test.
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        yield Path(tmpdir)


@pytest.fixture
def temp_file(temp_dir: Path) -> Generator[Path, None, None]:
    """Create a temporary file for testing.

    Args:
        temp_dir: Temporary directory fixture.

    Yields:
        Path: Path to the temporary file.
    """
    temp_file_path = temp_dir / "test_file.txt"
    temp_file_path.touch()
    yield temp_file_path


@pytest.fixture
def sample_csv_content() -> str:
    """Provide sample CSV content for testing.

    Returns:
        str: Sample CSV content with headers and data rows.
    """
    return "Name,Age,City\nJohn,30,New York\nJane,25,Los Angeles\n"


@pytest.fixture
def sample_xlsx_path() -> Path:
    """Provide path to a sample XLSX file from the test directory.

    Returns:
        Path: Path to a sample XLSX test file.
    """
    test_data_dir = Path(__file__).parent.parent / "test"
    return test_data_dir / "datetime.xlsx"


@pytest.fixture
def test_data_dir() -> Path:
    """Provide path to the test data directory.

    Returns:
        Path: Path to the directory containing test XLSX and CSV files.
    """
    return Path(__file__).parent.parent / "test"


@pytest.fixture(autouse=True)
def reset_environment() -> Generator[None, None, None]:
    """Reset environment variables after each test.

    This fixture automatically runs for every test and ensures that any
    environment variable changes made during a test don't affect other tests.
    """
    original_env = os.environ.copy()
    yield
    os.environ.clear()
    os.environ.update(original_env)


@pytest.fixture
def mock_config(mocker) -> dict:
    """Provide a mock configuration dictionary.

    Args:
        mocker: pytest-mock fixture for creating mocks.

    Returns:
        dict: Mock configuration dictionary with common settings.
    """
    return {
        "delimiter": ",",
        "encoding": "utf-8",
        "sheet_name": None,
        "skip_empty_lines": False,
    }


@pytest.fixture
def capture_output(mocker):
    """Fixture to capture stdout and stderr output.

    Args:
        mocker: pytest-mock fixture for creating mocks.

    Returns:
        Mock object that captures output.
    """
    import io
    output = io.StringIO()
    mocker.patch("sys.stdout", output)
    return output
