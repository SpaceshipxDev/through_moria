import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from PIL import Image as PILImage
import os
from io import BytesIO

def extract_customer_excel(file_path, images_output_dir):
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active

    # 1. Known schema (modify as needed)
    fallback_headers = ["序号", "图号", "图片", "名称", "材质", "数量", "表面处理", "备注"]

    # 2. Build actual headers, fallback if missing
    headers = []
    for idx, cell in enumerate(ws[1]):
        if cell.value is not None:
            headers.append(cell.value)
        else:
            headers.append(f"Unnamed_Column_{idx+1}")

    os.makedirs(images_output_dir, exist_ok=True)

    images_by_row = {}
    for img in ws._images:
        anchor_row = img.anchor._from.row + 1
        images_by_row[anchor_row] = img

    structured_data = []
    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        row_data = {}
        for col_idx, cell in enumerate(row):
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

    return structured_data

if __name__ == "__main__":
    file_path = "20250528-探野T4-3D打印手板加工清单.xlsx"
    images_dir = "extracted_images"

    data_list = extract_customer_excel(file_path, images_dir)

    import json
    print(json.dumps(data_list, indent=4, ensure_ascii=False))