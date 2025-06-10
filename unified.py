#!/usr/bin/env python3
"""
Excel Quote Generator MVP
Drop in Excel file → Get formatted quote Excel out
"""

import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from PIL import Image as PILImage
import os
import json
import sys
from io import BytesIO
from datetime import datetime
from openai import OpenAI

def extract_customer_excel(file_path, images_output_dir):
    """Extract data and images from customer Excel file"""
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active

    # Build headers
    headers = []
    for idx, cell in enumerate(ws[1]):
        if cell.value is not None:
            headers.append(cell.value)
        else:
            headers.append(f"Unnamed_Column_{idx+1}")

    os.makedirs(images_output_dir, exist_ok=True)

    # Get worksheet dimensions
    max_data_row = ws.max_row
    
    print("=== EXTRACTING IMAGES ===")
    
    # First pass: calculate all image positions and scores for each row
    image_candidates = {}  # {image_index: [(row, score), ...]}
    
    for i, img in enumerate(ws._images):
        anchor = img.anchor
        from_row = anchor._from.row + 1  # Convert to 1-based
        from_col = anchor._from.col + 1
        
        # Handle different anchor types
        if hasattr(anchor, 'to') and anchor.to:
            to_row = anchor.to.row + 1
            to_col = anchor.to.col + 1
        else:
            # Use image size to estimate end position
            to_row = from_row + 2  # Assume 2 rows high by default
            to_col = from_col + 1
        
        print(f"Image {i}: from=({from_row},{from_col}) to=({to_row},{to_col})")
        
        # Calculate scores for all possible data rows
        candidates = []
        for data_row in range(2, max_data_row + 1):
            score = calculate_image_row_score(from_row, to_row, data_row)
            if score > 0:  # Only consider positive scores
                candidates.append((data_row, score))
        
        # Sort by score (highest first)
        candidates.sort(key=lambda x: x[1], reverse=True)
        image_candidates[i] = candidates

    # Second pass: use Hungarian-like assignment to avoid conflicts
    images_by_row = {}
    assigned_images = set()
    
    # Start with images that have fewer good options (more constrained)
    image_constraint_order = sorted(
        image_candidates.keys(), 
        key=lambda i: len([c for c in image_candidates[i] if c[1] > 5])
    )
    
    for img_idx in image_constraint_order:
        if img_idx in assigned_images:
            continue
            
        candidates = image_candidates[img_idx]
        
        # Find best unassigned row
        for row, score in candidates:
            if row not in images_by_row:
                images_by_row[row] = ws._images[img_idx]
                assigned_images.add(img_idx)
                print(f"Assigned image {img_idx} to row {row} (score: {score:.1f})")
                break
        else:
            # No unassigned row found, handle conflict
            if candidates:
                best_row, best_score = candidates[0]
                if best_row in images_by_row:
                    # Check if we should replace the existing assignment
                    current_img_idx = list(ws._images).index(images_by_row[best_row])
                    current_candidates = image_candidates.get(current_img_idx, [])
                    current_score = next((s for r, s in current_candidates if r == best_row), 0)
                    
                    if best_score > current_score * 1.2:  # 20% better to replace
                        # Try to relocate the current image
                        relocated = False
                        for alt_row, alt_score in current_candidates:
                            if alt_row not in images_by_row and alt_score > current_score * 0.7:  # Accept 30% score drop
                                images_by_row[alt_row] = images_by_row[best_row]
                                relocated = True
                                break
                        
                        if relocated or len(current_candidates) <= 1:  # Replace if relocated or current has no alternatives
                            images_by_row[best_row] = ws._images[img_idx]
                            assigned_images.add(img_idx)

    # Third pass: handle any remaining unassigned images
    unassigned_images = set(range(len(ws._images))) - assigned_images
    unassigned_rows = set(range(2, max_data_row + 1)) - set(images_by_row.keys())
    
    if unassigned_images and unassigned_rows:
        print(f"Assigning remaining {len(unassigned_images)} images to {len(unassigned_rows)} rows")
        
        # Simple proximity-based assignment for remainders
        for img_idx in unassigned_images:
            if not unassigned_rows:
                break
                
            img = ws._images[img_idx]
            anchor = img.anchor
            from_row = anchor._from.row + 1
            
            # Find closest unassigned row
            closest_row = min(unassigned_rows, key=lambda r: abs(r - from_row))
            images_by_row[closest_row] = img
            unassigned_rows.remove(closest_row)

    # Build the structured data
    structured_data = []
    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        row_data = {}
        for col_idx, cell in enumerate(row):
            if col_idx < len(headers):
                column_name = headers[col_idx]
                row_data[column_name] = cell.value if cell.value is not None else "null"

        image_filename = None
        if row_idx in images_by_row:
            openpyxl_img: OpenpyxlImage = images_by_row[row_idx]
            image_filename = f"image_row_{row_idx}.png"
            image_path = os.path.join(images_output_dir, image_filename)
            img_bytes = openpyxl_img._data()
            img_pil = PILImage.open(BytesIO(img_bytes))
            img_pil.save(image_path)

        row_data['image_file'] = image_filename if image_filename is not None else "null"
        structured_data.append(row_data)

    print(f"✅ Extracted {len(structured_data)} rows with {len(images_by_row)} images")
    return structured_data

def calculate_image_row_score(img_from_row, img_to_row, data_row):
    """Calculate how well an image matches a data row"""
    
    # Ensure img_to_row is at least img_from_row
    img_to_row = max(img_to_row, img_from_row)
    
    # Calculate overlap between image span and data row
    overlap_start = max(img_from_row, data_row)
    overlap_end = min(img_to_row, data_row + 1)
    overlap = max(0, overlap_end - overlap_start)
    
    if overlap <= 0:
        # No overlap, but calculate proximity penalty
        if data_row < img_from_row:
            distance = img_from_row - data_row
        else:
            distance = data_row - img_to_row
        
        # Proximity score (decreases with distance)
        if distance <= 2:
            return max(0, 2 - distance)  # Score 2 for adjacent, 1 for 1 row away
        else:
            return 0
    
    # Base score from overlap amount
    overlap_score = overlap * 10
    
    # Bonus for perfect alignment (image starts exactly at data row)
    alignment_bonus = 3 if img_from_row == data_row else 0
    
    # Bonus for image center being close to row center
    img_center = (img_from_row + img_to_row) / 2
    row_center = data_row + 0.5
    center_distance = abs(img_center - row_center)
    center_bonus = max(0, 2 - center_distance)  # Up to 2 points for perfect centering
    
    # Penalty for very large images (they might span multiple rows incorrectly)
    img_height = img_to_row - img_from_row
    size_penalty = min(2, max(0, img_height - 3))  # Penalty for images > 3 rows tall
    
    total_score = overlap_score + alignment_bonus + center_bonus - size_penalty
    
    return max(0, total_score)

def process_with_qwen(extracted_data):
    """Process extracted data using Qwen API"""
    
    # Check for API key
    if not os.environ.get("DASHSCOPE_API_KEY"):
        print("❌ DASHSCOPE_API_KEY environment variable not set")
        return None
    
    # Initialize OpenAI client for Qwen
    client = OpenAI(
        api_key=os.environ.get("DASHSCOPE_API_KEY"),
        base_url="https://dashscope.aliyuncs.com/compatible-mode/v1",
    )
    
    # Create prompt
    prompt = f"""You must translate and restructure the extracted JSON data into our internal CNC machining company Excel log format:

    Internal format required:
    - Serial_Number: 
    - Part_Name: 
    - Quantity: 
    - Material: 
    - Machining_Process: 
    - Surface_Finish: 
    - Notes: (or N/A if none)
    - image_file: (exact file name or null if not present - DO NOT modify filenames)

    Instructions to follow:
    - Translate Chinese keys exactly according to my mapping.
    - Keep original data values the same.
    - DON'T modify or lose the "image_file" reference. Pass it along unchanged.
    - Output structured valid JSON ONLY. NO OTHER TEXT. Don't output empty roles. 

    Here is the input JSON data:

    {json.dumps(extracted_data, ensure_ascii=False)}"""

    try:
        print("🤖 Processing with Qwen API...")
        response = client.chat.completions.create(
            model="qwen-turbo",
            messages=[{"role": "user", "content": prompt}],
        )
        
        # Get the response content
        qwen_response = response.choices[0].message.content.strip()
        
        # Clean and parse the response
        if qwen_response.startswith('```json'):
            qwen_response = qwen_response[7:-3]
        elif qwen_response.startswith('```'):
            qwen_response = qwen_response[3:-3]
        
        # Parse JSON response
        processed_data = json.loads(qwen_response)
        print(f"✅ Processed {len(processed_data)} items with Qwen")
        return processed_data
        
    except json.JSONDecodeError as e:
        print(f"❌ Error parsing JSON: {e}")
        print(f"Raw response: {qwen_response}")
        return None
    except Exception as e:
        print(f"❌ API call failed: {e}")
        return None

def generate_quote_excel(processed_data, output_filename):
    """Generate formatted quote Excel file"""
    
    # Create workbook and worksheet
    wb = Workbook()
    ws = wb.active
    ws.title = "手板报价"

    # Company header information
    current_date = datetime.now().strftime("%Y-%m-%d")

    # Set up the header section (rows 1-12)
    ws.merge_cells('A1:I1')  # Updated to span 9 columns
    ws['A1'] = f"手板报价"
    ws['A1'].font = Font(name='SimSun', size=16, bold=True)
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 25

    # Company information section
    ws.merge_cells('A2:D2')
    ws['A2'] = " "
    ws['A2'].font = Font(name='SimSun', size=10)
    ws['A2'].alignment = Alignment(horizontal='left', vertical='center')

    ws.merge_cells('E2:I2')  # Updated to span to column I
    ws['E2'] = "乙方:杭州越依模型科技有限公司"
    ws['E2'].font = Font(name='SimSun', size=10)
    ws['E2'].alignment = Alignment(horizontal='left', vertical='center')

    # Contact information
    contact_info = [
        ("联系人:", "联系人:傅士勤"),
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
        
        ws.merge_cells(f'E{i}:I{i}')  # Updated to span to column I
        ws[f'E{i}'] = right
        ws[f'E{i}'].font = Font(name='SimSun', size=9)
        ws[f'E{i}'].alignment = Alignment(horizontal='left', vertical='center')

    # Set row heights for header section
    for row in range(2, 8):
        ws.row_dimensions[row].height = 18

    # REMOVED: Product specification header section (手板类型, 手板精度, 备注)

    # Main table headers (row 8) - Updated row number since we removed the ugly section
    table_headers = ["序号", "零件图片", "零件名", "表面", "材质", "数量", "单价(未税)", "总价(未税)", "备注"]
    header_widths = [6, 12, 15, 8, 10, 8, 12, 12, 15]

    for col_num, (header, width) in enumerate(zip(table_headers, header_widths), 1):
        cell = ws.cell(row=8, column=col_num)  # Changed from row 10 to row 8
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

    ws.row_dimensions[8].height = 25  # Changed from row 10 to row 8

    # Add data rows
    data_start_row = 9  # Changed from 11 to 9
    
    for idx, row_data in enumerate(processed_data, start=data_start_row):
        # Set row height for images
        ws.row_dimensions[idx].height = 60
        
        # Extract surface finish with proper null handling
        surface_finish = row_data.get("Surface_Finish", "")
        if surface_finish is None or surface_finish == "null":
            surface_display = ""
        else:
            surface_display = str(surface_finish).replace("120#", "").replace("+", "+\n")
        
        # Extract quantity - keep as is, let Excel handle the math
        quantity = row_data.get("Quantity", 0)
        try:
            if isinstance(quantity, str) and quantity.strip():
                quantity = float(quantity)
            elif not isinstance(quantity, (int, float)):
                quantity = 0
        except (ValueError, TypeError):
            quantity = 0
        
        # Data to write (no more Python calculations)
        row_values = [
            row_data.get("Serial_Number", ""),
            "",  # Image placeholder
            row_data.get("Part_Name", ""),
            surface_display,
            row_data.get("Material", ""),
            quantity,  # Column F
            0,  # Unit price placeholder - Column G
            None,  # Total price will be formula - Column H
            row_data.get("Notes", "") if row_data.get("Notes") not in ["null", None, "N/A"] else ""
        ]
        
        # Write data with formatting
        for col_num, value in enumerate(row_values, 1):
            cell = ws.cell(row=idx, column=col_num)
            
            # Special handling for total price column (Column H = 8)
            if col_num == 8:  # Total price column
                # Set formula: 总价 = 数量 * 单价 (F * G)
                cell.value = f"=F{idx}*G{idx}"
            else:
                cell.value = value
            
            cell.font = Font(name='SimSun', size=9)
            
            # Center alignment for specific columns
            if col_num in [1, 2, 4, 5, 6, 7, 8]:  # Serial, Image, Surface, Material, Quantity, Unit Price, Total Price
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
                    print(f"⚠️  Error adding image for row {idx}: {e}")

    # Add totals row with SUM formula
    total_row = len(processed_data) + data_start_row
    ws.merge_cells(f'A{total_row}:G{total_row}')  # Merge from A to G
    ws[f'A{total_row}'] = "合计:"
    ws[f'A{total_row}'].font = Font(name='SimSun', size=10, bold=True)
    ws[f'A{total_row}'].alignment = Alignment(horizontal="right", vertical="center")

    # Add SUM formula for total amount in Column H
    first_data_row = data_start_row
    last_data_row = len(processed_data) + data_start_row - 1
    ws[f'H{total_row}'] = f"=SUM(H{first_data_row}:H{last_data_row})"
    ws[f'H{total_row}'].font = Font(name='SimSun', size=10, bold=True)
    ws[f'H{total_row}'].alignment = Alignment(horizontal="center", vertical="center")

    # Add borders to total row
    for col in range(1, 10):  # Updated to cover 9 columns
        cell = ws.cell(row=total_row, column=col)
        cell.border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )

    # Footer information - now references the total cell
    footer_start = total_row + 2
    footer_info = [
        f"未 税 总 价: =H{total_row} (人民币)",  # Reference the total cell
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
            # For the total price display, we'll show it as text with formula reference
            if "=H" in info:
                ws[f'A{row_num}'] = f"未 税 总 价: (人民币)"  # Keep as text, user can see total in the table
            else:
                ws[f'A{row_num}'] = info
            ws[f'A{row_num}'].font = Font(name='SimSun', size=10)
            ws[f'A{row_num}'].alignment = Alignment(horizontal='left', vertical='center')
        else:
            ws.merge_cells(f'A{row_num}:I{row_num}')  # Updated to span to column I
            ws[f'A{row_num}'] = info
            ws[f'A{row_num}'].font = Font(name='SimSun', size=9)
            ws[f'A{row_num}'].alignment = Alignment(horizontal='left', vertical='center')

    # Signature section
    signature_row = footer_start + len(footer_info) + 2
    ws.merge_cells(f'G{signature_row}:I{signature_row}')  # Updated to use columns G-I
    ws[f'G{signature_row}'] = "乙方签名确认"
    ws[f'G{signature_row}'].font = Font(name='SimSun', size=10, bold=True)
    ws[f'G{signature_row}'].alignment = Alignment(horizontal='center', vertical='center')

    ws.merge_cells(f'G{signature_row + 1}:I{signature_row + 1}')  # Updated to use columns G-I
    ws[f'G{signature_row + 1}'] = f"{current_date}"
    ws[f'G{signature_row + 1}'].font = Font(name='SimSun', size=10)
    ws[f'G{signature_row + 1}'].alignment = Alignment(horizontal='center', vertical='center')

    # Save the workbook
    wb.save(output_filename)
    print(f"✅ Quote Excel generated: {output_filename}")
    return output_filename

def main():
    """Main pipeline function"""
    
    # Check command line arguments
    if len(sys.argv) < 2:
        print("Usage: python quote_generator.py <input_excel_file>")
        print("Example: python quote_generator.py input.xlsx")
        sys.exit(1)
    
    input_file = sys.argv[1]
    
    # Check if input file exists
    if not os.path.exists(input_file):
        print(f"❌ Input file not found: {input_file}")
        sys.exit(1)
    
    print(f"🚀 Starting Excel Quote Generator")
    print(f"📁 Input file: {input_file}")
    
    try:
        # Step 1: Extract data and images from customer Excel
        print("\n=== STEP 1: EXTRACTING DATA ===")
        extracted_data = extract_customer_excel(input_file, "extracted_images")
        
        if not extracted_data:
            print("❌ No data extracted from Excel file")
            sys.exit(1)
        
        # Step 2: Process with Qwen API
        print("\n=== STEP 2: PROCESSING WITH QWEN ===")
        processed_data = process_with_qwen(extracted_data)
        
        if not processed_data:
            print("❌ Failed to process data with Qwen API")
            sys.exit(1)
        
        # Step 3: Generate quote Excel
        print("\n=== STEP 3: GENERATING QUOTE ===")
        current_date = datetime.now().strftime("%Y-%m-%d")
        output_filename = f"手板报价单_{current_date}.xlsx"
        
        quote_file = generate_quote_excel(processed_data, output_filename)
        
        print(f"\n🎉 SUCCESS!")
        print(f"📊 Processed {len(processed_data)} items")
        print(f"📝 Quote generated: {quote_file}")
        
    except Exception as e:
        print(f"❌ Pipeline failed: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()