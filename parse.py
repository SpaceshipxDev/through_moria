from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import json
import os

# Load processed data from think.py
try:
    with open('gemini_processed_data.json', 'r', encoding='utf-8') as f:
        gemini_output_data = json.load(f)
except FileNotFoundError:
    print("❌ gemini_processed_data.json not found. Run think.py first.")
    exit(1)

# Create workbook and worksheet
wb = Workbook()
ws = wb.active
ws.title = "手板报价"

# Company header information
current_date = datetime.now().strftime("%Y-%m-%d")
quote_number = f"YNMX-25-{datetime.now().strftime('%m-%d')}-314"

# Set up the header section (rows 1-12)
ws.merge_cells('A1:H1')
ws['A1'] = f"手板报价  编号:{quote_number}"
ws['A1'].font = Font(name='SimSun', size=16, bold=True)
ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
ws.row_dimensions[1].height = 25

# Company information section
ws.merge_cells('A2:D2')
ws['A2'] = "甲方:杭州微信软件有限公司"
ws['A2'].font = Font(name='SimSun', size=10)
ws['A2'].alignment = Alignment(horizontal='left', vertical='center')

ws.merge_cells('E2:H2')
ws['E2'] = "乙方:杭州越依模型科技有限公司"
ws['E2'].font = Font(name='SimSun', size=10)
ws['E2'].alignment = Alignment(horizontal='left', vertical='center')

# Contact information
contact_info = [
    ("联系人:舒康洪", "联系人:傅士勤"),
    ("TEL:", "TEL: 13777479066"),
    ("FAX:", "FAX:"),
    ("E-mail:", "E-mail:"),
    ("地址:", "地址:杭州市富阳区东洲工业功能区1号路11号")
]

for i, (left, right) in enumerate(contact_info, start=3):
    ws.merge_cells(f'A{i}:D{i}')
    ws[f'A{i}'] = left
    ws[f'A{i}'].font = Font(name='SimSun', size=9)
    ws[f'A{i}'].alignment = Alignment(horizontal='left', vertical='center')
    
    ws.merge_cells(f'E{i}:H{i}')
    ws[f'E{i}'] = right
    ws[f'E{i}'].font = Font(name='SimSun', size=9)
    ws[f'E{i}'].alignment = Alignment(horizontal='left', vertical='center')

# Set row heights for header section
for row in range(2, 8):
    ws.row_dimensions[row].height = 18

# Product specification header (row 8)
ws.merge_cells('A8:C8')
ws['A8'] = "手板类型"
ws['A8'].font = Font(name='SimSun', size=10, bold=True)
ws['A8'].alignment = Alignment(horizontal='center', vertical='center')
ws['A8'].fill = PatternFill(start_color="E6E6E6", end_color="E6E6E6", fill_type="solid")

ws.merge_cells('D8:F8')
ws['D8'] = "手板精度"
ws['D8'].font = Font(name='SimSun', size=10, bold=True)
ws['D8'].alignment = Alignment(horizontal='center', vertical='center')
ws['D8'].fill = PatternFill(start_color="E6E6E6", end_color="E6E6E6", fill_type="solid")

ws.merge_cells('G8:H8')
ws['G8'] = "备注"
ws['G8'].font = Font(name='SimSun', size=10, bold=True)
ws['G8'].alignment = Alignment(horizontal='center', vertical='center')
ws['G8'].fill = PatternFill(start_color="E6E6E6", end_color="E6E6E6", fill_type="solid")

# Add borders to header section
for row in range(8, 9):
    for col in range(1, 9):
        cell = ws.cell(row=row, column=col)
        cell.border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )

# Main table headers (row 10)
table_headers = ["序号", "零件图片", "零件名", "表面", "材质", "数量", "价(未税)", "备注"]
header_widths = [6, 12, 15, 8, 10, 8, 12, 15]

for col_num, (header, width) in enumerate(zip(table_headers, header_widths), 1):
    cell = ws.cell(row=10, column=col_num)
    cell.value = header
    cell.font = Font(name='SimSun', size=10, bold=True)
    cell.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )
    
    # Set column width
    ws.column_dimensions[get_column_letter(col_num)].width = width

ws.row_dimensions[10].height = 25

# Add data rows
for idx, row_data in enumerate(gemini_output_data, start=11):
    # Set row height for images
    ws.row_dimensions[idx].height = 60
    
    # Extract surface finish (remove numbers and special chars for display)
    surface_display = row_data.get("Surface_Finish", "").replace("120#", "").replace("+", "+\n")
    
    # Data to write
    row_values = [
        row_data.get("Serial_Number", ""),
        "",  # Image placeholder
        row_data.get("Part_Name", ""),
        surface_display,
        row_data.get("Material", ""),
        row_data.get("Quantity", ""),
        0,  # Price placeholder
        row_data.get("Notes", "") if row_data.get("Notes") not in ["null", None, "N/A"] else ""
    ]
    
    # Write data with formatting
    for col_num, value in enumerate(row_values, 1):
        cell = ws.cell(row=idx, column=col_num)
        cell.value = value
        cell.font = Font(name='SimSun', size=9)
        
        # Center alignment for specific columns
        if col_num in [1, 2, 4, 5, 6, 7]:  # Serial, Image, Surface, Material, Quantity, Price
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        
        # Add borders
        cell.border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )

    # Handle image insertion
    if row_data.get("image_file") and row_data["image_file"] != "null":
        image_path = os.path.join("extracted_images", row_data["image_file"])
        if os.path.exists(image_path):
            try:
                img = XLImage(image_path)
                # Resize image to fit in cell
                max_size = 50
                if img.width > max_size or img.height > max_size:
                    ratio = min(max_size/img.width, max_size/img.height)
                    img.width = int(img.width * ratio)
                    img.height = int(img.height * ratio)
                
                # Position image in the cell
                img.anchor = f"B{idx}"
                ws.add_image(img)
                
            except Exception as e:
                print(f"Error adding image for row {idx}: {e}")

# Add totals row
total_row = len(gemini_output_data) + 11
ws.merge_cells(f'A{total_row}:F{total_row}')
ws[f'A{total_row}'] = "合计:"
ws[f'A{total_row}'].font = Font(name='SimSun', size=10, bold=True)
ws[f'A{total_row}'].alignment = Alignment(horizontal="right", vertical="center")

ws[f'G{total_row}'] = 0
ws[f'G{total_row}'].font = Font(name='SimSun', size=10, bold=True)
ws[f'G{total_row}'].alignment = Alignment(horizontal="center", vertical="center")

# Add borders to total row
for col in range(1, 9):
    cell = ws.cell(row=total_row, column=col)
    cell.border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

# Footer information
footer_start = total_row + 2
footer_info = [
    f"未 税 总 价: (人民币)",
    "手板加工周期:",
    "付款方式: 月结30天",
    "交货日期: 确认后 (7) 日内完成",
    "验收标准: 依据甲方2D、3D、说明文档等相关文件进行验收",
    "备 注:",
    "此报价单适用于所有杭州海康威视科技有限公司的子公司及关联公司。",
    "双方以含税价格结算，具体税率按国家税务政策规定，供应商需提供合格的增值税发票，否则按基"
]

for i, info in enumerate(footer_info):
    row_num = footer_start + i
    if info.startswith("未 税"):
        ws.merge_cells(f'A{row_num}:F{row_num}')
        ws[f'A{row_num}'] = info
        ws[f'A{row_num}'].font = Font(name='SimSun', size=10)
        ws[f'A{row_num}'].alignment = Alignment(horizontal='left', vertical='center')
    else:
        ws.merge_cells(f'A{row_num}:H{row_num}')
        ws[f'A{row_num}'] = info
        ws[f'A{row_num}'].font = Font(name='SimSun', size=9)
        ws[f'A{row_num}'].alignment = Alignment(horizontal='left', vertical='center')

# Signature section
signature_row = footer_start + len(footer_info) + 2
ws.merge_cells(f'F{signature_row}:H{signature_row}')
ws[f'F{signature_row}'] = "乙方签名确认"
ws[f'F{signature_row}'].font = Font(name='SimSun', size=10, bold=True)
ws[f'F{signature_row}'].alignment = Alignment(horizontal='center', vertical='center')

ws.merge_cells(f'F{signature_row + 1}:H{signature_row + 1}')
ws[f'F{signature_row + 1}'] = f"{current_date}"
ws[f'F{signature_row + 1}'].font = Font(name='SimSun', size=10)
ws[f'F{signature_row + 1}'].alignment = Alignment(horizontal='center', vertical='center')

# Save the workbook
output_filename = f"手板报价单_{current_date}.xlsx"
wb.save(output_filename)
print(f"✅ 专业中文报价单已创建: {output_filename}")
print(f"📊 包含 {len(gemini_output_data)} 个零件项目")