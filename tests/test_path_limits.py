import os

import pytest

import orders


def test_fit_filename_shortens_when_path_too_long(tmp_path):
    long_dir = tmp_path
    # Create nested directories to make the absolute path relatively long.
    for idx in range(4):
        long_dir = long_dir / ("segment" + str(idx))
        long_dir.mkdir()

    filename = "Standard bon_BOM-07_Powder coating _ Git zwart - RAL 9005_2025-10-26.xlsx"
    limit = len(os.path.abspath(long_dir)) + 40
    safe_name = orders._fit_filename_within_path(str(long_dir), filename, max_path=limit)

    assert safe_name.endswith(".xlsx")
    assert len(os.path.join(os.path.abspath(long_dir), safe_name)) <= limit
    assert safe_name != filename


def test_fit_filename_returns_original_when_within_limit(tmp_path):
    filename = "Bestelbon_PN1_2025-10-26.pdf"
    limit = len(os.path.abspath(tmp_path)) + len(filename) + 10
    safe_name = orders._fit_filename_within_path(str(tmp_path), filename, max_path=limit)

    assert safe_name == filename


def test_fit_filename_raises_when_directory_too_long(tmp_path):
    # Choose a limit that makes the directory itself exceed the limit.
    limit = len(os.path.abspath(tmp_path))
    with pytest.raises(OSError):
        orders._fit_filename_within_path(str(tmp_path), "example.pdf", max_path=limit)
