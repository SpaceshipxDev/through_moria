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
        # Strategy: prioritize by image center and size
        image_height = max(to_row_1based - from_row_1based, 1)
        image_center = (from_row_1based + to_row_1based) / 2
        
        best_row = None
        best_score = -1
        
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
            
            # Score based on multiple factors:
            # 1. Overlap amount (most important)
            # 2. Distance from image center to row center
            # 3. Bonus if image anchor is in this row
            
            if overlap > 0:
                center_distance = abs(image_center - (data_row + 0.5))
                anchor_bonus = 2 if from_row_1based == data_row else 0
                
                # Higher score is better
                score = overlap * 10 - center_distance + anchor_bonus
                
                if score > best_score:
                    best_score = score
                    best_row = data_row
        
        # Alternative: if no good overlap, use proximity to anchor
        if best_score <= 0:
            min_distance = float('inf')
            for data_row in range(2, max_data_row + 1):
                distance = abs(from_row_1based - data_row)
                if distance < min_distance:
                    min_distance = distance
                    best_row = data_row
            best_score = -min_distance  # Negative score for fallback cases
        
        print(f"Image {i}: height={image_height}, center={image_center:.1f}, score={best_score:.1f}")
        
        if best_row:
            # Check if this row already has an image assigned
            if best_row not in images_by_row:
                images_by_row[best_row] = img
                print(f"  -> Mapped to data row {best_row}")
            else:
                # Conflict resolution: compare scores
                print(f"  -> Conflict with row {best_row}!")
                
                # Find the current image in this row and compare
                current_img = images_by_row[best_row]
                current_img_idx = list(ws._images).index(current_img)
                
                # Calculate score for current image
                curr_anchor = current_img.anchor
                curr_from_row = curr_anchor._from.row + 1
                curr_to_row = (curr_anchor.to.row + 1) if hasattr(curr_anchor, 'to') and curr_anchor.to else curr_from_row
                curr_center = (curr_from_row + curr_to_row) / 2
                curr_center_distance = abs(curr_center - (best_row + 0.5))
                curr_anchor_bonus = 2 if curr_from_row == best_row else 0
                
                # Calculate overlap for current image
                curr_overlap_start = max(curr_from_row, best_row)
                curr_overlap_end = min(max(curr_to_row, curr_from_row + 1), best_row + 1)
                curr_overlap = max(0, curr_overlap_end - curr_overlap_start)
                
                curr_score = curr_overlap * 10 - curr_center_distance + curr_anchor_bonus
                
                print(f"    Current image {current_img_idx} score: {curr_score:.1f}")
                print(f"    New image {i} score: {best_score:.1f}")
                
                if best_score > curr_score:
                    # New image wins, reassign current image
                    print(f"    New image wins! Reassigning current image...")
                    
                    # Find alternative row for current image
                    for alt_row in range(2, max_data_row + 1):
                        if alt_row not in images_by_row or alt_row == best_row:
                            if alt_row != best_row:  # Don't reassign to same row
                                images_by_row[alt_row] = current_img
                                print(f"    Moved current image to row {alt_row}")
                                break
                    
                    # Assign new image to contested row
                    images_by_row[best_row] = img
                    print(f"  -> Assigned new image to row {best_row}")
                else:
                    # Current image stays, find alternative for new image
                    print(f"    Current image stays, finding alternative...")
                    for alt_row in range(2, max_data_row + 1):
                        if alt_row not in images_by_row:
                            images_by_row[alt_row] = img
                            print(f"  -> Assigned to alternative row {alt_row}")
                            break
                    else:
                        print(f"  -> No alternative row found, image not mapped")
        else:
            print(f"  -> No suitable row found")

    print(f"\n=== FINAL MAPPING ===")
    for row, img in images_by_row.items():
        print(f"Row {row} -> Image")

    # Build the structured data
    structured_data = []
    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        row_data = {}
        for col_idx, cell in enumerate(row):
            if col_idx < len(headers):
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