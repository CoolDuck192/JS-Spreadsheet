#!/usr/bin/env python3
"""Generate deterministic openpyxl-authored XLSX import fixtures."""

from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from tempfile import NamedTemporaryFile
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference


ROOT = Path(__file__).resolve().parents[1]
OUTPUT_DIRECTORY = ROOT / "src" / "test" / "fixtures" / "xlsx"
FIXED_DATETIME = datetime(2026, 1, 1, tzinfo=timezone.utc)
FIXED_ZIP_TIME = (2026, 1, 1, 0, 0, 0)


def base_workbook() -> tuple[Workbook, object]:
    workbook = Workbook()
    workbook.properties.created = FIXED_DATETIME
    workbook.properties.modified = FIXED_DATETIME
    worksheet = workbook.active
    worksheet.title = "S"
    worksheet.append(["Q", "V"])
    for quarter, value in (("Q1", 120), ("Q2", 150), ("Q3", 90), ("Q4", 200)):
        worksheet.append([quarter, value])
    return workbook, worksheet


def save_deterministic(workbook: Workbook, name: str) -> None:
    OUTPUT_DIRECTORY.mkdir(parents=True, exist_ok=True)
    destination = OUTPUT_DIRECTORY / name
    with NamedTemporaryFile(suffix=".xlsx", dir="/tmp") as temporary:
        workbook.save(temporary.name)
        with ZipFile(temporary.name, "r") as source, ZipFile(
            destination, "w", compression=ZIP_DEFLATED, compresslevel=9
        ) as output:
            for source_info in source.infolist():
                target_info = ZipInfo(source_info.filename, FIXED_ZIP_TIME)
                target_info.compress_type = ZIP_DEFLATED
                target_info.external_attr = source_info.external_attr
                target_info.create_system = source_info.create_system
                output.writestr(target_info, source.read(source_info.filename))


def generate_chart_fixture() -> None:
    workbook, worksheet = base_workbook()
    chart = BarChart()
    chart.add_data(Reference(worksheet, min_col=2, min_row=1, max_row=5), titles_from_data=True)
    chart.set_categories(Reference(worksheet, min_col=1, min_row=2, max_row=5))
    worksheet.add_chart(chart, "D2")
    save_deterministic(workbook, "variant-chart.xlsx")


def main() -> None:
    generate_chart_fixture()


if __name__ == "__main__":
    main()
