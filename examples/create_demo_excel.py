import pandas as pd
from pathlib import Path

# Sheet 1: Line Items
sheet1 = pd.DataFrame([
    ["Item No.", "Description", "Unit", "Quantity", "Unit Price", "Total Price"],
    ["1.1", "Excavation for foundation", "m³", 150.00, 25.00, 3750.00],
    ["1.2", "Concrete foundation", "m³", 75.00, 120.00, 9000.00],
    ["", "", "", "", "Subtotal", 12750.00],
    ["2.1", "Brickwork", "m²", 200.00, 45.00, 9000.00],
    ["", "", "", "", "Total", 21750.00],
])

# Sheet 2: Project Info + Table
sheet2 = pd.DataFrame([
    ["Project Information", "", "", "", "", ""],
    ["Project Name:", "Sample Building Project", "", "", "", ""],
    ["Client:", "ABC Construction Ltd", "", "", "", ""],
    ["Date:", "2024-01-15", "", "", "", ""],
    ["", "", "", "", "", ""],
    ["Item", "Specification", "Qty", "Rate", "Amount", "Remarks"],
    [1, "Site preparation", 1, 5000.00, 5000.00, ""],
    [2, "Foundation work", 1, 15000.00, 15000.00, ""],
    ["", "", "", "", "Total", 20000.00],
])

# Sheet 3: Notes
sheet3 = pd.DataFrame([
    ["Notes and Conditions", "", "", "", "", ""],
    ["1. All prices are inclusive of taxes", "", "", "", "", ""],
    ["2. Payment terms: 30 days", "", "", "", "", ""],
    ["3. Delivery: As per schedule", "", "", "", "", ""],
    ["", "", "", "", "", ""],
    ["Contact Information", "", "", "", "", ""],
    ["Phone:", "+1-234-567-8900", "", "", "", ""],
    ["Email:", "info@sample.com", "", "", "", ""],
])

# Write to Excel
output_path = Path(__file__).parent / "sample_boq.xlsx"
with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    sheet1.to_excel(writer, sheet_name="Sheet1", index=False, header=False)
    sheet2.to_excel(writer, sheet_name="Sheet2", index=False, header=False)
    sheet3.to_excel(writer, sheet_name="Sheet3", index=False, header=False)

print(f"Demo Excel file created at: {output_path}") 