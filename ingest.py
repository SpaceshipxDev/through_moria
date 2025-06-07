import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from PIL import Image as PILImage
import os
from io import BytesIO
import json

def extract_customer_excel(file_path, images_output_dir):
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active

    # Create images output directory if it doesn't exist
    os.makedirs(images_output_dir, exist_ok=True)

    # Map images directly to their anchored row number
    images_by_row = {}
    for img in ws._images:
        anchor_row = img.anchor._from.row + 1  # Excel rows start at 1
        images_by_row[anchor_row] = img

    structured_data = []

    # Process rows without considering headers
    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        # Extract just cell values as a simple list
        cell_values = [cell.value for cell in row]

        # Extract and save the corresponding image (if exists)
        image_filename = None
        if row_idx in images_by_row:
            openpyxl_img: OpenpyxlImage = images_by_row[row_idx]
            image_filename = f"image_row_{row_idx}.png"
            image_path = os.path.join(images_output_dir, image_filename)
            img_bytes = openpyxl_img._data()
            img_pil = PILImage.open(BytesIO(img_bytes))
            img_pil.save(image_path)

        # Append clearly structured data row
        structured_data.append({
            "row_number": row_idx,
            "cells": cell_values,
            "image_file": image_filename
        })

    return structured_data

if __name__ == "__main__":
    file_path = "20250528-探野T4-3D打印手板加工清单.xlsx"
    images_dir = "extracted_images"

    data_list = extract_customer_excel(file_path, images_dir)

    # Checking the structured simplified output
    print(json.dumps(data_list, indent=4, ensure_ascii=False))