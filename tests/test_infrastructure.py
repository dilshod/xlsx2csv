"""Validation tests for the testing infrastructure setup.

This module contains basic tests to verify that the testing infrastructure
is properly configured and working correctly.
"""
import sys
from pathlib import Path

import pytest


class TestInfrastructure:
    """Tests to validate the testing infrastructure."""

    def test_pytest_is_working(self):
        """Verify that pytest is properly installed and working."""
        assert True

    def test_python_version(self):
        """Verify that Python version is 3.7 or higher."""
        assert sys.version_info >= (3, 7), "Python 3.7+ is required"

    def test_project_root_exists(self):
        """Verify that we can access the project root directory."""
        project_root = Path(__file__).parent.parent
        assert project_root.exists()
        assert (project_root / "pyproject.toml").exists()

    def test_main_module_exists(self):
        """Verify that the main xlsx2csv module exists."""
        project_root = Path(__file__).parent.parent
        xlsx2csv_path = project_root / "xlsx2csv.py"
        assert xlsx2csv_path.exists(), "xlsx2csv.py should exist"

    @pytest.mark.unit
    def test_unit_marker(self):
        """Verify that the 'unit' marker is configured."""
        assert True

    @pytest.mark.integration
    def test_integration_marker(self):
        """Verify that the 'integration' marker is configured."""
        assert True

    @pytest.mark.slow
    def test_slow_marker(self):
        """Verify that the 'slow' marker is configured."""
        assert True


class TestFixtures:
    """Tests to validate that shared fixtures are working."""

    def test_temp_dir_fixture(self, temp_dir):
        """Verify that temp_dir fixture creates a directory."""
        assert temp_dir.exists()
        assert temp_dir.is_dir()

    def test_temp_file_fixture(self, temp_file):
        """Verify that temp_file fixture creates a file."""
        assert temp_file.exists()
        assert temp_file.is_file()

    def test_sample_csv_content_fixture(self, sample_csv_content):
        """Verify that sample_csv_content fixture provides data."""
        assert isinstance(sample_csv_content, str)
        assert len(sample_csv_content) > 0
        assert "Name,Age,City" in sample_csv_content

    def test_test_data_dir_fixture(self, test_data_dir):
        """Verify that test_data_dir fixture points to test directory."""
        assert test_data_dir.exists()
        assert test_data_dir.is_dir()
        assert (test_data_dir / "datetime.xlsx").exists()

    def test_mock_config_fixture(self, mock_config):
        """Verify that mock_config fixture provides configuration."""
        assert isinstance(mock_config, dict)
        assert "delimiter" in mock_config
        assert "encoding" in mock_config


class TestMocking:
    """Tests to validate pytest-mock functionality."""

    def test_mocker_fixture_available(self, mocker):
        """Verify that pytest-mock is available."""
        mock_obj = mocker.Mock()
        mock_obj.test_method.return_value = "mocked"
        assert mock_obj.test_method() == "mocked"

    def test_patch_functionality(self, mocker):
        """Verify that mocker.patch works correctly."""
        mocker.patch("os.getcwd", return_value="/fake/path")
        import os
        assert os.getcwd() == "/fake/path"
