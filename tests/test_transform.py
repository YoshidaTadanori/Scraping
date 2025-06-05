import os
import csv
import pathlib
import sys

# Ensure the repository root is on sys.path
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))

from Scraping import transform_and_save_data


def test_transform_and_save_data_creates_expected_csv(tmp_path):
    data = [
        ['2025/06/01', 'Hotel A', 100],
        ['2025/06/01', 'Hotel B', 200],
        ['2025/06/02', 'Hotel A', 150],
        ['2025/06/02', 'Hotel C', 0],
    ]

    output_file = tmp_path / "out.csv"
    transform_and_save_data(data, str(output_file))

    with open(output_file, newline='', encoding='utf-8') as f:
        rows = list(csv.reader(f))

    expected = [
        ['\u65e5\u4ed8', 'Hotel A', 'Hotel B', 'Hotel C', '\u5e73\u5747'],
        ['2025/06/01', '100', '200', '0', '150'],
        ['2025/06/02', '150', '0', '0', '150'],
    ]
    assert rows == expected
