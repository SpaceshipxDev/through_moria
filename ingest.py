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

    # Get worksheet dimensions
    max_data_row = ws.max_row
    
    # More precise image mapping using cell dimensions
    images_by_row = {}
    
    # Get row heights to better understand positioning
    row_heights = {}
    for row_num in range(1, max_data_row + 1):
        row_heights[row_num] = ws.row_dimensions[row_num].height or 15  # Default Excel row height
    
    print("=== IMAGE MAPPING DEBUG ===")
    for i, img in enumerate(ws._images):
        # Get precise anchor position
        anchor = img.anchor
        from_row = anchor._from.row  # 0-based
        from_col = anchor._from.col  # 0-based
        
        # Get the "to" position (bottom-right of image)
        to_row = anchor.to.row if hasattr(anchor, 'to') and anchor.to else from_row
        to_col = anchor.to.col if hasattr(anchor, 'to') and anchor.to else from_col
        
        # Convert to 1-based Excel row numbers
        from_row_1based = from_row + 1
        to_row_1based = to_row + 1
        
        print(f"Image {i}: from_row={from_row_1based}, to_row={to_row_1based}, from_col={from_col}, to_col={to_col}")
        
        # Find which data row this image most likely belongs to
        # Strategy: find the data row that has the most overlap with the image
        best_row = None
        best_overlap = 0
        
        for data_row in range(2, max_data_row + 1):  # Data starts from row 2
            # Calculate overlap between image and this data row
            image_start = from_row_1based
            image_end = max(to_row_1based, from_row_1based + 1)  # At least 1 row high
            
            row_start = data_row
            row_end = data_row + 1
            
            # Calculate overlap
            overlap_start = max(image_start, row_start)
            overlap_end = min(image_end, row_end)
            overlap = max(0, overlap_end - overlap_start)
            
            if overlap > best_overlap:
                best_overlap = overlap
                best_row = data_row
        
        # Alternative: if no good overlap, use proximity
        if best_overlap == 0:
            min_distance = float('inf')
            for data_row in range(2, max_data_row + 1):
                # Distance from image center to row center
                image_center = (from_row_1based + to_row_1based) / 2
                distance = abs(image_center - data_row)
                if distance < min_distance:
                    min_distance = distance
                    best_row = data_row
        
        if best_row:
            # Only assign if this row doesn't already have an image
            if best_row not in images_by_row:
                images_by_row[best_row] = img
                print(f"  -> Mapped to data row {best_row} (overlap: {best_overlap})")
            else:
                print(f"  -> Row {best_row} already has image, trying next best...")
                # Find next best row that's available
                alternatives = []
                for data_row in range(2, max_data_row + 1):
                    if data_row not in images_by_row:
                        distance = abs(from_row_1based - data_row)
                        alternatives.append((distance, data_row))
                
                if alternatives:
                    alternatives.sort()
                    chosen_row = alternatives[0][1]
                    images_by_row[chosen_row] = img
                    print(f"  -> Assigned to alternative row {chosen_row}")
                else:
                    print(f"  -> No available rows, image not mapped")

    print(f"\n=== FINAL MAPPING ===")
    for row, img in images_by_row.items():
        print(f"Row {row} -> Image")

    # Build the structured data
    structured_data = []
    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        row_data = {}
        
    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        row_data = {}
        for col_idx, cell in enumerate(row):
            if col_idx < len(headers):
                column_name = headers[col_idx]
                # Patch: Nulls become "null"
                row_data[column_name] = cell.value if cell.value is not None else "null"

        image_filename = None
        if row_idx in images_by_row:
            openpyxl_img: OpenpyxlImage = images_by_row[row_idx]
            image_filename = f"image_row_{row_idx}.png"
            image_path = os.path.join(images_output_dir, image_filename)
            img_bytes = openpyxl_img._data()
            img_pil = PILImage.open(BytesIO(img_bytes))
            img_pil.save(image_path)

        # Patch: Nulls become "null"
        row_data['image_file'] = image_filename if image_filename is not None else "null"
        structured_data.append(row_data)

    # Final debug info
    print(f"\n=== SUMMARY ===")
    print(f"Total images: {len(ws._images)}")
    print(f"Total data rows: {len(structured_data)}")
    print(f"Images mapped: {len(images_by_row)}")
    
    rows_without_images = [i for i, row_data in enumerate(structured_data, start=2) if row_data.get('image_file') is None]
    if rows_without_images:
        print(f"Rows without images: {rows_without_images}")
    
    return structured_data

if __name__ == "__main__":
    file_path = "data/20250528-探野T4-3D打印手板加工清单.xlsx"
    images_dir = "extracted_images"

    data_list = extract_customer_excel(file_path, images_dir)

    import json
    print(json.dumps(data_list, indent=4, ensure_ascii=False))