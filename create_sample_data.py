"""
Generate a sample items_data.xlsx file with product data.
Run this once to create the Excel data source.
"""
import openpyxl

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Items"

# Headers
headers = ["Item Name", "Size", "Unit Price", "Discount"]
ws.append(headers)

# Sample product data: (name, packing size, unit price, discount amount)
items = [
    ("Paracetamol 500mg", "10 Tablets", 25.00, 2.00),
    ("Amoxicillin 250mg", "15 Capsules", 85.00, 5.00),
    ("Vitamin C 1000mg", "20 Tablets", 120.00, 10.00),
    ("Cough Syrup", "100 ml", 65.00, 3.00),
    ("Ibuprofen 400mg", "10 Tablets", 30.00, 2.50),
    ("Antacid Gel", "170 ml", 95.00, 8.00),
    ("Cetirizine 10mg", "10 Tablets", 35.00, 3.00),
    ("Multivitamin", "30 Tablets", 250.00, 20.00),
    ("Bandage Roll", "1 Roll", 40.00, 0.00),
    ("Antiseptic Cream", "15 gm", 55.00, 4.00),
    ("ORS Powder", "5 Sachets", 30.00, 2.00),
    ("Pain Relief Spray", "50 ml", 180.00, 15.00),
    ("Eye Drops", "10 ml", 70.00, 5.00),
    ("Hand Sanitizer", "200 ml", 99.00, 10.00),
    ("Face Mask", "10 Pack", 50.00, 5.00),
    ("Digital Thermometer", "1 Unit", 350.00, 25.00),
    ("Cotton Roll", "100 gm", 45.00, 0.00),
    ("Calcium Tablets", "30 Tablets", 190.00, 15.00),
    ("Iron Supplement", "20 Tablets", 110.00, 8.00),
    ("Protein Powder", "500 gm", 650.00, 50.00),
]

for item in items:
    ws.append(item)

# Auto-adjust column widths
for col in ws.columns:
    max_length = max(len(str(cell.value or "")) for cell in col)
    ws.column_dimensions[col[0].column_letter].width = max_length + 4

wb.save("items_data.xlsx")
print("Created items_data.xlsx with", len(items), "items.")
