from .generator_utils import create_workbook, save_workbook

def generate_department():

    departments = [
        ["FIN","Finance"],
        ["HR","Human Resources"],
        ["IT","Information Technology"],
        ["OPS","Operations"],
        ["SAL","Sales"],
        ["MKT","Marketing"],
        ["SCM","Supply Chain"],
        ["LEG","Legal"],
        ["ADM","Administration"],
        ["CS","Customer Success"]
    ]

    wb, ws = create_workbook("Department")

    ws.append([
        "DepartmentCode",
        "DepartmentName"
    ])

    for department in departments:
        ws.append(department)

    save_workbook(
        wb,
        "Data/Master/Department.xlsx"
    )
