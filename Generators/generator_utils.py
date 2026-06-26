from pathlib import Path
from openpyxl import Workbook


def create_workbook(sheet_name):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    return wb, ws


def save_workbook(workbook, relative_path):
    project_root = Path(__file__).resolve().parent.parent
    output_path = project_root / relative_path
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    print(f"✅ Saved: {output_path}")
