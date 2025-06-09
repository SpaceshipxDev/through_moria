import json
from openai import OpenAI

# Use your DeepSeek API key here
client = OpenAI(
    api_key="sk-d4c80c767d914bdc8dad316f9ec6100c",
    base_url="https://api.deepseek.com"
)

your_step1_data = [...]  # your data as above

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

response = client.chat.completions.create(
    model="deepseek-chat",
    messages=[
        {"role": "user", "content": prompt},
    ],
    stream=False
)

ai_output = response.choices[0].message.content.strip()

if ai_output.startswith('```json'):
    ai_output = ai_output[7:-3]
elif ai_output.startswith('```'):
    ai_output = ai_output[3:-3]

try:
    ai_data = json.loads(ai_output)
    print("✅ Successfully parsed DeepSeek response")
    print(f"📊 Processed {len(ai_data)} items")
    with open('deepseek_processed_data.json', 'w', encoding='utf-8') as f:
        json.dump(ai_data, f, ensure_ascii=False, indent=2)
    print("💾 Saved processed data to deepseek_processed_data.json")
except json.JSONDecodeError as e:
    print(f"❌ Error parsing JSON: {e}")
    print(f"Raw response: {ai_output}")
    ai_data = None