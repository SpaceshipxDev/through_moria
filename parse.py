from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
import json
import os

# Gemini structured data you've received (saved here clearly)
gemini_output_data = [
{
    "Serial_Number": 22,
    "Part_Name": "显示屏壳体",
    "Quantity": 1,
    "Material": "树脂",
    "Machining_Process": "N/A",
    "Surface_Finish": "无",
    "Notes": "null",
    "image_file": "null"
  },
  {
    "Serial_Number": 23,
    "Part_Name": "右饰片",
    "Quantity": 1,
    "Material": "树脂",
    "Machining_Process": "N/A",
    "Surface_Finish": "无",
    "Notes": "null",
    "image_file": "image_row_25.png"
  },
  {
    "Serial_Number": 24,
    "Part_Name": "左饰片",
    "Quantity": 1,
    "Material": "树脂",
    "Machining_Process": "N/A",
    "Surface_Finish": "无",
    "Notes": "null",
    "image_file": "image_row_26.png"
  },
  {
    "Serial_Number": 25,
    "Part_Name": "装饰片1",
    "Quantity": 1,
    "Material": "树脂",
    "Machining_Process": "N/A",
    "Surface_Finish": "无",
    "Notes": "null",
    "image_file": "image_row_27.png"
  },
  {
    "Serial_Number": 26,
    "Part_Name": "装饰片2",
    "Quantity": 1,
    "Material": "树脂",
    "Machining_Process": "N/A",
    "Surface_Finish": "无",
    "Notes": "null",
    "image_file": "image_row_28.png"
  },
  {
    "Serial_Number": 27,
    "Part_Name": "转接座",
    "Quantity": 1,
    "Material": "树脂",
    "Machining_Process": "N/A",
    "Surface_Finish": "无",
    "Notes": "null",
    "image_file": "image_row_29.png"
  },
  {
    "Serial_Number": 28,
    "Part_Name": "转轴座",
    "Quantity": 1,
    "Material": "树脂",
    "Machining_Process": "N/A",
    "Surface_Finish": "无",
    "Notes": "null",
    "image_file": "image_row_30.png"
  },
  {
    "Serial_Number": 29,
    "Part_Name": "主板",
    "Quantity": 1,
    "Material": "树脂",
    "Machining_Process": "N/A",
    "Surface_Finish": "无",
    "Notes": "null",
    "image_file": "image_row_31.png"
  },
  {
    "Serial_Number": 30,
    "Part_Name": "主壳体",
    "Quantity": 1,
    "Material": "树脂",
    "Machining_Process": "N/A",
    "Surface_Finish": "无",
    "Notes": "null",
    "image_file": "image_row_32.png"
  }
  # Add other records...
]

# Step (1) clearly create Excel workbook
wb = Workbook()
ws = wb.active
ws.title = "CNC_Log"

# (2) Write headers clearly
headers = ["Serial_Number", "Part_Name", "Quantity", "Material",
           "Machining_Process", "Surface_Finish", "Notes", "Image"]

# Write headers to Excel
ws.append(headers)

# (3) Write your extracted structured data clearly to Excel
for idx, row_data in enumerate(gemini_output_data, start=2):  # start at excel row 2 (row 1 headers)
    
    # adding the textual data clearly
    row = [
        row_data["Serial_Number"],
        row_data["Part_Name"],
        row_data["Quantity"],
        row_data["Material"],
        row_data["Machining_Process"],
        row_data["Surface_Finish"],
        "" if row_data["Notes"] in ["null", None] else row_data["Notes"],
        ""  # Image cell placeholder; embedded separately
    ]
    
    ws.append(row)

    # add images clearly
    if row_data["image_file"] and os.path.exists(row_data["image_file"]):
        img = XLImage(row_data["image_file"])

        # Adjust cell width and height clearly (optional but recommended)
        ws.row_dimensions[idx].height = 80  # Adjust clearly to fit images nicely
        ws.column_dimensions["H"].width = 30  # "H" is the image column in this case

        # Embed clearly & visually into Excel (column "H")
        img.anchor = f"H{idx}"  # Position clearly cells in column H
        ws.add_image(img)
    else:
        print(f"No valid image: Skipped embedding image at Excel row {idx}")

# (4) Clearly save final Excel file
wb.save("final_CNC_log.xlsx")
print("✅ Clearly structured CNC Excel generated successfully.")