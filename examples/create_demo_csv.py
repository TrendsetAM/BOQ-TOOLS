import csv
from pathlib import Path

# Create CSV data
csv_data = [
    ["Item No.", "Description", "Unit", "Quantity", "Unit Price", "Total Price"],
    ["1.1", "Excavation for foundation", "m³", "150.00", "25.00", "3750.00"],
    ["1.2", "Concrete foundation", "m³", "75.00", "120.00", "9000.00"],
    ["", "", "", "", "Subtotal", "12750.00"],
    ["2.1", "Brickwork", "m²", "200.00", "45.00", "9000.00"],
    ["", "", "", "", "Total", "21750.00"],
]

# Save as CSV
output_path = Path(__file__).parent / "sample_boq.csv"
with open(output_path, 'w', newline='', encoding='utf-8') as csvfile:
    writer = csv.writer(csvfile)
    writer.writerows(csv_data)

print(f"Demo CSV file created at: {output_path}")
print("Note: This is a CSV file for testing. For full Excel functionality, install pandas and openpyxl.") 