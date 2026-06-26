from openpyxl import Workbook

companies = [
    ["NS001", "Nova Manufacturing", "India", "Manufacturing", "INR"],
    ["NS002", "Nova Retail", "India", "Retail", "INR"]
]

wb = Workbook()
ws = wb.active

for company in companies:
    ws.append(company)

wb.save("companies.xlsx")