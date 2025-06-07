import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from PIL import Image as PILImage
import os
from io import BytesIO

def extract_customer_excel(file_path, images_output_dir):
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active

    # Step 1: Build headers list (for correct mapping)
    headers = [cell.value if cell.value else f"Column_{i+1}" for i, cell in enumerate(ws[1])]

    # Create images directory
    os.makedirs(images_output_dir, exist_ok=True)

    # Map image positions to row indices (1-based)
    images_by_row = {}
    for img in ws._images:
        anchor_row = img.anchor._from.row + 1  # openpyxl is zero-indexed
        images_by_row[anchor_row] = img

    # Collect structured data
    structured_data = []
    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):  # skip header row
        row_data = {}
        for col_idx, cell in enumerate(row):
            column_name = headers[col_idx]
            row_data[column_name] = cell.value

        # Attach image if present
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

    return structured_data

if __name__ == "__main__":
    file_path = "20250528-探野T4-3D打印手板加工清单.xlsx"
    images_dir = "extracted_images"

    data_list = extract_customer_excel(file_path, images_dir)

    import json
    print(json.dumps(data_list, indent=4, ensure_ascii=False))  # Add ensure_ascii=False for readable Chinese