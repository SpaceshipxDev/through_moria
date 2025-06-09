# Updated Gemini API example code clearly & simply.
import json
from google import genai
from google.genai import types

# Your API Key (Replace obviously with your actual key)
client = genai.Client(api_key="AIzaSyCuq4sZ5x_covMuJEuJ6vHbs0I2fKpAQDM")

# Extracted data from step 1 clearly loaded here
your_step1_data = [
    {
        "序号": 1,
        "零件名称": 190169051,
        "数量": 6,
        "材质": "6061AL",
        "工艺": "机加",
        "外观处理": "120#喷砂+黑色氧化",
        "备注": "null",
        "image_file": "null"
    },
    {
        "序号": 2,
        "零件名称": "da-14-010578",
        "数量": 3,
        "材质": "6061AL",
        "工艺": "机加",
        "外观处理": "120#喷砂+黑色氧化",
        "备注": "局部镭雕去氧化面",
        "image_file": "null"
    },
    {
        "序号": 3,
        "零件名称": "da-14-010579",
        "数量": 3,
        "材质": "6061AL",
        "工艺": "机加",
        "外观处理": "120#喷砂+黑色氧化",
        "备注": "局部镭雕去氧化面",
        "image_file": "null"
    },
    {
        "序号": 4,
        "零件名称": "da-14-010592",
        "数量": 12,
        "材质": "6061AL",
        "工艺": "机加",
        "外观处理": "120#喷砂+黑色氧化",
        "备注": "局部镭雕去氧化面",
        "image_file": "null"
    },
    {
        "序号": 5,
        "零件名称": "da-14-010634",
        "数量": 3,
        "材质": "6061AL",
        "工艺": "机加",
        "外观处理": "120#喷砂+黑色氧化",
        "备注": "局部镭雕去氧化面",
        "image_file": "null"
    },
    {
        "序号": 6,
        "零件名称": "xieyijiexi",
        "数量": 3,
        "材质": "6061AL",
        "工艺": "机加",
        "外观处理": "120#喷砂+黑色氧化",
        "备注": "局部镭雕去氧化面",
        "image_file": "null"
    }
]

# clearly create your prompt to Gemini:
prompt = f"""
You must translate and restructure the extracted JSON data into our internal CNC machining company Excel log format:

Internal format clearly required:
- Serial_Number:
- Part_Name:
- Quantity:
- Material:
- Machining_Process:
- Surface_Finish:
- Notes: (or N/A if none)
- image_file: (exact file name or null if not present clearly DO NOT modify filenames)

Instructions clearly to follow:
- Clearly translate Chinese keys exactly according to my mapping.
- Keep original data values the same.
- DON'T modify or lose the "image_file" reference. Clearly pass it along unchanged.
- Output clearly structured valid JSON ONLY. NO OTHER TEXT.

Here is the input JSON data clearly:

{json.dumps(your_step1_data, ensure_ascii=False)}
"""

# clearly formatted Gemini API Call
response = client.models.generate_content(
    model="gemini-2.5-flash-preview-05-20",
    config=types.GenerateContentConfig(
        thinking_config = types.ThinkingConfig(
            thinking_budget=0,
        )),
    contents=prompt
)

# Clean and parse the response
gemini_response = response.text.strip()
# Remove any markdown formatting if present
if gemini_response.startswith('```json'):
    gemini_response = gemini_response[7:-3]
elif gemini_response.startswith('```'):
    gemini_response = gemini_response[3:-3]

# Parse JSON response
try:
    gemini_output_data = json.loads(gemini_response)
    print("✅ Successfully parsed Gemini response")
    print(f"📊 Processed {len(gemini_output_data)} items")
    
    # Save the processed data for the Excel parser
    with open('gemini_processed_data.json', 'w', encoding='utf-8') as f:
        json.dump(gemini_output_data, f, ensure_ascii=False, indent=2)
    
    print("💾 Saved processed data to gemini_processed_data.json")
    
except json.JSONDecodeError as e:
    print(f"❌ Error parsing JSON: {e}")
    print(f"Raw response: {gemini_response}")
    gemini_output_data = None