# Updated to use Volcengine Ark API instead of Gemini
import json
import os
from volcenginesdkarkruntime import Ark

# Initialize Ark client
client = Ark(api_key=os.environ.get("ARK_API_KEY"))

# Extracted data from step 1
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

# Create your prompt
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
- Output structured valid JSON ONLY. NO OTHER TEXT.

Here is the input JSON data:

{json.dumps(your_step1_data, ensure_ascii=False)}"""

# Make API call using Ark
try:
    response = client.chat.completions.create(
        model="doubao-1-5-pro-32k-250115",
        messages=[{"content": prompt, "role": "user"}],
    )
    
    # Print reasoning content if available (for deep thinking models)
    if hasattr(response.choices[0].message, 'reasoning_content'):
        print("🧠 Model reasoning:")
        print(response.choices[0].message.reasoning_content)
        print("-" * 50)
    
    # Get the response content
    ark_response = response.choices[0].message.content.strip()
    
    # Clean and parse the response
    if ark_response.startswith('```json'):
        ark_response = ark_response[7:-3]
    elif ark_response.startswith('```'):
        ark_response = ark_response[3:-3]
    
    # Parse JSON response
    ark_output_data = json.loads(ark_response)
    print("✅ Successfully parsed Ark response")
    print(f"📊 Processed {len(ark_output_data)} items")
    
    # Save the processed data for the Excel parser
    with open('ark_processed_data.json', 'w', encoding='utf-8') as f:
        json.dump(ark_output_data, f, ensure_ascii=False, indent=2)
    
    print("💾 Saved processed data to ark_processed_data.json")

except json.JSONDecodeError as e:
    print(f"❌ Error parsing JSON: {e}")
    print(f"Raw response: {ark_response}")
    ark_output_data = None
except Exception as e:
    print(f"❌ API call failed: {e}")
    ark_output_data = None