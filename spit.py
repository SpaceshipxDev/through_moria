import os
from openpyxl import Workbook
from openpyxl.drawing.image import Image as OpenpyxlImage

# Define the directory clearly containing your images
images_dir = "extracted_images"

# Output Excel file name clearly defined
output_excel = "output.xlsx"

# Create Workbook and worksheet
wb = Workbook()
ws = wb.active
ws.title = "Components Log"

# Define the column headers clearly:
headers = ["Component ID", "Description", "Material", "Image"]
ws.append(headers)

# Simulate some fake data (just for this quick embedding test clearly):
# later the real textual information will come directly from your LLM structured output

data_for_testing = [
    {"id": "CMP001", "desc": "Component 1", "material": "Steel", "image_filename": "component1.png"},
    {"id": "CMP002", "desc": "Component 2", "material": "Aluminum", "image_filename": "component2.png"},
    {"id": "CMP003", "desc": "Component 3", "material": "Titanium", "image_filename": "component3.png"},
]

# Start clearly embedding:
row_num = 2  # Start clearly at row 2 (row 1 clearly header row)

for item in data_for_testing:
    ws.cell(row=row_num, column=1, value=item["id"])
    ws.cell(row=row_num, column=2, value=item["desc"])
    ws.cell(row=row_num, column=3, value=item["material"])

    # Construct path clearly and embed image into Excel
    image_path = os.path.join(images_dir, item["image_filename"])

    if os.path.exists(image_path):
        img = OpenpyxlImage(image_path)

        # Optional: clearly resize images to fit Excel cells nicely if they are too big:
        img.width = 100  # Optional clearly defined width
        img.height = 100  # Optional clearly defined height

        # Clearly indicate the cell location (column D, current row)
        img_anchor = f'D{row_num}'
        ws.add_image(img, img_anchor)
    else:
        print(f"Image file not found clearly: {image_path}")

    row_num += 1

# Clearly save workbook into your output file
wb.save(output_excel)

print("✅ Excel file with embedded images generated successfully!")