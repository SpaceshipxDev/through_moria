import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from PIL import Image as PILImage
import os
from io import BytesIO

def extract_customer_excel(file_path, images_output_dir):
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

    # Get total number of data rows
    max_data_row = ws.max_row

    # Method 1: Simple closest row mapping
    images_by_row = {}
    
    # Alternative Method 2: More precise mapping using both row and column info
    # Uncomment this section if Method 1 doesn't work well
    """
    for img in ws._images:
        anchor_row = img.anchor._from.row
        anchor_col = img.anchor._from.col
        
        # Check if image overlaps with any data row
        for data_row_idx in range(1, max_data_row):  # 0-based iteration
            actual_row_num = data_row_idx + 2  # Convert to actual row number (starting from row 2)
            
            # Check if image row position is close to this data row
            row_diff = abs(anchor_row - data_row_idx)
            if row_diff <= 1:  # Allow 1 row tolerance
                if actual_row_num not in images_by_row:
                    images_by_row[actual_row_num] = img
                    print(f"Image at ({anchor_row}, {anchor_col}) mapped to data row {actual_row_num}")
                break
    """
    
    # Method 1: Simple approach - map to closest row
    for img in ws._images:
        anchor_row_0based = img.anchor._from.row
        anchor_row_1based = anchor_row_0based + 1
        
        # Find closest data row
        closest_data_row = None
        min_distance = float('inf')
        
        for data_row in range(2, max_data_row + 1):
            distance = abs(anchor_row_1based - data_row)
            if distance < min_distance:
                min_distance = distance
                closest_data_row = data_row
        
        # Always map to closest row
        if closest_data_row:
            if closest_data_row not in images_by_row:
                images_by_row[closest_data_row] = img
                print(f"Image anchored at row {anchor_row_1based} mapped to data row {closest_data_row} (distance: {min_distance})")
            else:
                print(f"Row {closest_data_row} already has an image, skipping image at row {anchor_row_1based}")
    
    # Second pass: try to assign remaining images to rows without images
    unmapped_images = []
    for img in ws._images:
        anchor_row_1based = img.anchor._from.row + 1
        is_mapped = any(mapped_img == img for mapped_img in images_by_row.values())
        if not is_mapped:
            unmapped_images.append((img, anchor_row_1based))
    
    if unmapped_images:
        print(f"Found {len(unmapped_images)} unmapped images")
        # Try to assign them to rows that don't have images yet
        for img, original_row in unmapped_images:
            for data_row in range(2, max_data_row + 1):
                if data_row not in images_by_row:
                    images_by_row[data_row] = img
                    print(f"Assigned orphan image (was at row {original_row}) to empty data row {data_row}")
                    break

    structured_data = []
    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        row_data = {}
        for col_idx, cell in enumerate(row):
            if col_idx < len(headers):  # Safety check
                column_name = headers[col_idx]
                row_data[column_name] = cell.value

        image_filename = None
        if row_idx in images_by_row:
            openpyxl_img: OpenpyxlImage = images_by_row[row_idx]
            image_filename = f"image_row_{row_idx}.png"
            image_path = os.path.join(images_output_dir, image_filename)
            img_bytes = openpyxl_img._data()
            img_pil = PILImage.open(BytesIO(img_bytes))
            img_pil.save(image_path)

        row_data['image_file'] = image_filename
        structured_data.append(row_data)

    # Debug info
    print(f"Found {len(ws._images)} images total")
    print(f"Found {len(structured_data)} data rows")
    print(f"Mapped {len(images_by_row)} images to data rows")
    
    # Show which rows don't have images
    rows_without_images = []
    for i, row_data in enumerate(structured_data, start=2):
        if row_data.get('image_file') is None:
            rows_without_images.append(i)
    
    if rows_without_images:
        print(f"Rows without images: {rows_without_images}")
    
    return structured_data

if __name__ == "__main__":
    file_path = "data/20250528-探野T4-3D打印手板加工清单.xlsx"
    images_dir = "extracted_images"

    data_list = extract_customer_excel(file_path, images_dir)

    import json
    print(json.dumps(data_list, indent=4, ensure_ascii=False))