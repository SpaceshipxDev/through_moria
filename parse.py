from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import json
import os

# Gemini structured data
gemini_output_data = [
  {
    "Serial_Number": 1,
    "Part_Name": 190169051,
    "Quantity": 6,
    "Material": "6061AL",
    "Machining_Process": "机加",
    "Surface_Finish": "120#喷砂+黑色氧化",
    "Notes": "N/A",
    "image_file": "null"
  },
  {
    "Serial_Number": 2,
    "Part_Name": "da-14-010578",
    "Quantity": 3,
    "Material": "6061AL",
    "Machining_Process": "机加",
    "Surface_Finish": "120#喷砂+黑色氧化",
    "Notes": "局部镭雕去氧化面",
    "image_file": "null"
  },
  {
    "Serial_Number": 3,
    "Part_Name": "da-14-010579",
    "Quantity": 3,
    "Material": "6061AL",
    "Machining_Process": "机加",
    "Surface_Finish": "120#喷砂+黑色氧化",
    "Notes": "局部镭雕去氧化面",
    "image_file": "null"
  },
  {
    "Serial_Number": 4,
    "Part_Name": "da-14-010592",
    "Quantity": 12,
    "Material": "6061AL",
    "Machining_Process": "机加",
    "Surface_Finish": "120#喷砂+黑色氧化",
    "Notes": "局部镭雕去氧化面",
    "image_file": "null"
  },
  {
    "Serial_Number": 5,
    "Part_Name": "da-14-010634",
    "Quantity": 3,
    "Material": "6061AL",
    "Machining_Process": "机加",
    "Surface_Finish": "120#喷砂+黑色氧化",
    "Notes": "局部镭雕去氧化面",
    "image_file": "null"
  },
  {
    "Serial_Number": 6,
    "Part_Name": "xieyijiexi",
    "Quantity": 3,
    "Material": "6061AL",
    "Machining_Process": "机加",
    "Surface_Finish": "120#喷砂+黑色氧化",
    "Notes": "局部镭雕去氧化面",
    "image_file": "null"
  }
]

# Create workbook and worksheet
wb = Workbook()
ws = wb.active
ws.title = "CNC_Log"

# Define headers
headers = ["Serial Number", "Part Name", "Quantity", "Material",
           "Machining Process", "Surface Finish", "Notes", "Image"]

# Write headers with formatting
for col_num, header in enumerate(headers, 1):
    cell = ws.cell(row=1, column=col_num)
    cell.value = header
    cell.font = Font(bold=True, size=12, color="FFFFFF")
    cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

# Set column widths for better spacing
column_widths = {
    'A': 12,  # Serial Number
    'B': 20,  # Part Name
    'C': 10,  # Quantity
    'D': 15,  # Material
    'E': 18,  # Machining Process
    'F': 15,  # Surface Finish
    'G': 25,  # Notes
    'H': 35   # Image (wider for images)
}

for col_letter, width in column_widths.items():
    ws.column_dimensions[col_letter].width = width

# Add data rows with proper formatting
for idx, row_data in enumerate(gemini_output_data, start=2):
    # Set row height to accommodate images properly
    ws.row_dimensions[idx].height = 120
    
    # Data to write
    row_values = [
        row_data["Serial_Number"],
        row_data["Part_Name"],
        row_data["Quantity"],
        row_data["Material"],
        row_data["Machining_Process"],
        row_data["Surface_Finish"],
        "" if row_data["Notes"] in ["null", None] else row_data["Notes"],
        ""  # Image placeholder
    ]
    
    # Write data with formatting
    for col_num, value in enumerate(row_values, 1):
        cell = ws.cell(row=idx, column=col_num)
        cell.value = value
        cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        cell.font = Font(size=10)
        
        # Add borders
        cell.border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )
        
        # Alternate row coloring for better readability
        if idx % 2 == 0:
            cell.fill = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")

    # Handle image insertion with proper sizing and positioning
    if row_data["image_file"] and row_data["image_file"] != "null":
        image_path = os.path.join("extracted_images", row_data["image_file"])
        if os.path.exists(image_path):
            try:
                img = XLImage(image_path)
                
                # Resize image to fit well in cell with margins
                # Maximum dimensions while leaving margins
                max_width = 200  # pixels
                max_height = 100  # pixels
                
                # Calculate scaling to maintain aspect ratio
                if img.width > max_width or img.height > max_height:
                    width_ratio = max_width / img.width
                    height_ratio = max_height / img.height
                    ratio = min(width_ratio, height_ratio)
                    
                    img.width = int(img.width * ratio)
                    img.height = int(img.height * ratio)
                
                # Position image with some margin from cell edges
                # Calculate cell position with offset for margins
                col_letter = 'H'
                cell_address = f"{col_letter}{idx}"
                
                # Add some offset to center the image in the cell
                img.anchor = cell_address
                
                # Add image to worksheet
                ws.add_image(img)
                
            except Exception as e:
                print(f"Error adding image for row {idx}: {e}")
                # Add "No Image" text if image fails
                ws[f"H{idx}"].value = "No Image"
                ws[f"H{idx}"].alignment = Alignment(horizontal="center", vertical="center")
    else:
        # Add "No Image" for missing images
        ws[f"H{idx}"].value = "No Image Available"
        ws[f"H{idx}"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"H{idx}"].font = Font(italic=True, color="999999")

# Freeze the header row for better navigation
ws.freeze_panes = "A2"

# Auto-filter for the header row
ws.auto_filter.ref = ws.dimensions

# Save the workbook
wb.save("professional_CNC_log.xlsx")
print("✅ Professional CNC Excel file created successfully with proper formatting!")