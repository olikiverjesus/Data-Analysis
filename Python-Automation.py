import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

# Load data from CSV
purchases = pd.read_csv("purchases.csv")
usage = pd.read_csv("usage.csv")

# Merge data on 'material_id'
merged_data = pd.merge(purchases, usage, on='material_id', how='left')

# Calculate total waste cost and waste percentage
merged_data['total_waste_cost'] = merged_data['quantity_wasted'] * merged_data['unit_price']
merged_data['waste_percentage'] = (merged_data['quantity_wasted'] / merged_data['quantity_purchased']) * 100

# Create a new Excel file
wb = Workbook()
ws = wb.active
ws.title = "Material Report"

# Add headers
headers = ["Material ID", "Material Name", "Supplier", "Purchase Date", "Quantity Purchased", "Unit Price", "Total Cost", "Quantity Used", "Quantity Wasted", "Total Waste Cost", "Waste Percentage"]
ws.append(headers)

# Add formatted data
for row in merged_data.itertuples(index=False):
    ws.append(list(row))

# Adjust styles
for col in ws.iter_cols(min_col=1, max_col=len(headers), min_row=1, max_row=1):
    for cell in col:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')

# Save the file
wb.save("Material_Report.xlsx")

print("Excel report generated successfully!")
