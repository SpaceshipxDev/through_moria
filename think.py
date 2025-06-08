# Updated Gemini API example code clearly & simply.
import json
from google import genai
from google.genai import types

# Your API Key (Replace obviously with your actual key)
client = genai.Client(api_key="AIzaSyCuq4sZ5x_covMuJEuJ6vHbs0I2fKpAQDM")

# Extracted data from step 1 clearly loaded here
your_step1_data = [
    {
        "20250528-探野T4-3D打印手板加工清单": "序号",
        "Unnamed_Column_2": "图号",
        "Unnamed_Column_3": "图片",
        "Unnamed_Column_4": "名称",
        "Unnamed_Column_5": "材质",
        "Unnamed_Column_6": "数量",
        "Unnamed_Column_7": "表面处理",
        "Unnamed_Column_8": "备注",
        "image_file": "null"
    },
    {
        "20250528-探野T4-3D打印手板加工清单": 1,
        "Unnamed_Column_2": "电池仓",
        "Unnamed_Column_3": "null",
        "Unnamed_Column_4": "电池仓",
        "Unnamed_Column_5": "树脂",
        "Unnamed_Column_6": 1,
        "Unnamed_Column_7": "无",
        "Unnamed_Column_8": "null",
        "image_file": "image_row_3.png"
    },
    {
        "20250528-探野T4-3D打印手板加工清单": 2,
        "Unnamed_Column_2": "电池盖",
        "Unnamed_Column_3": "null",
        "Unnamed_Column_4": "电池盖",
        "Unnamed_Column_5": "树脂",
        "Unnamed_Column_6": 1,
        "Unnamed_Column_7": "无",
        "Unnamed_Column_8": "null",
        "image_file": "image_row_4.png"
    },
    {
        "20250528-探野T4-3D打印手板加工清单": 3,
        "Unnamed_Column_2": "前壳",
        "Unnamed_Column_3": "null",
        "Unnamed_Column_4": "前壳",
        "Unnamed_Column_5": "树脂",
        "Unnamed_Column_6": 1,
        "Unnamed_Column_7": "无",
        "Unnamed_Column_8": "null",
        "image_file": "image_row_5.png"
    },
    {
        "20250528-探野T4-3D打印手板加工清单": 4,
        "Unnamed_Column_2": "按键1",
        "Unnamed_Column_3": "null",
        "Unnamed_Column_4": "按键1",
        "Unnamed_Column_5": "树脂",
        "Unnamed_Column_6": 1,
        "Unnamed_Column_7": "无",
        "Unnamed_Column_8": "null",
        "image_file": "image_row_6.png"
    },
    {
        "20250528-探野T4-3D打印手板加工清单": 5,
        "Unnamed_Column_2": "按键2",
        "Unnamed_Column_3": "null",
        "Unnamed_Column_4": "按键2",
        "Unnamed_Column_5": "树脂",
        "Unnamed_Column_6": 1,
        "Unnamed_Column_7": "无",
        "Unnamed_Column_8": "null",
        "image_file": "image_row_7.png"
    },
    {
        "20250528-探野T4-3D打印手板加工清单": 6,
        "Unnamed_Column_2": "按键3",
        "Unnamed_Column_3": "null",
        "Unnamed_Column_4": "按键3",
        "Unnamed_Column_5": "树脂",
        "Unnamed_Column_6": 1,
        "Unnamed_Column_7": "无",
        "Unnamed_Column_8": "null",
        "image_file": "image_row_8.png"
    },
    {
        "20250528-探野T4-3D打印手板加工清单": 7,
        "Unnamed_Column_2": "按键板",
        "Unnamed_Column_3": "null",
        "Unnamed_Column_4": "按键板",
        "Unnamed_Column_5": "树脂",
        "Unnamed_Column_6": 1,
        "Unnamed_Column_7": "无",
        "Unnamed_Column_8": "null",
        "image_file": "image_row_9.png"
    },
    {
        "20250528-探野T4-3D打印手板加工清单": 8,
        "Unnamed_Column_2": "电源板",
        "Unnamed_Column_3": "null",
        "Unnamed_Column_4": "电源板",
        "Unnamed_Column_5": "树脂",
        "Unnamed_Column_6": 1,
        "Unnamed_Column_7": "无",
        "Unnamed_Column_8": "null",
        "image_file": "image_row_10.png"
    },
    {
        "20250528-探野T4-3D打印手板加工清单": 9,
        "Unnamed_Column_2": "负极簧片",
        "Unnamed_Column_3": "null",
        "Unnamed_Column_4": "负极簧片",
        "Unnamed_Column_5": "树脂",
        "Unnamed_Column_6": 1,
        "Unnamed_Column_7": "无",
        "Unnamed_Column_8": "null",
        "image_file": "image_row_11.png"
    },
    {
        "20250528-探野T4-3D打印手板加工清单": 10,
        "Unnamed_Column_2": "负极接线片",
        "Unnamed_Column_3": "null",
        "Unnamed_Column_4": "负极接线片",
        "Unnamed_Column_5": "树脂",
        "Unnamed_Column_6": 1,
        "Unnamed_Column_7": "无",
        "Unnamed_Column_8": "null",
        "image_file": "image_row_12.png"
    },
    {
        "20250528-探野T4-3D打印手板加工清单": 11,
        "Unnamed_Column_2": "接口板",
        "Unnamed_Column_3": "null",
        "Unnamed_Column_4": "接口板",
        "Unnamed_Column_5": "树脂",
        "Unnamed_Column_6": 1,
        "Unnamed_Column_7": "无",
        "Unnamed_Column_8": "null",
        "image_file": "image_row_13.png"
    },
    {
        "20250528-探野T4-3D打印手板加工清单": 12,
        "Unnamed_Column_2": "激光装饰片",
        "Unnamed_Column_3": "null",
        "Unnamed_Column_4": "激光装饰片",
        "Unnamed_Column_5": "树脂",
        "Unnamed_Column_6": 1,
        "Unnamed_Column_7": "无",
        "Unnamed_Column_8": "null",
        "image_file": "image_row_14.png"
    },
    {
        "20250528-探野T4-3D打印手板加工清单": 13,
        "Unnamed_Column_2": "镜头",
        "Unnamed_Column_3": "null",
        "Unnamed_Column_4": "镜头",
        "Unnamed_Column_5": "树脂",
        "Unnamed_Column_6": 1,
        "Unnamed_Column_7": "无",
        "Unnamed_Column_8": "null",
        "image_file": "image_row_15.png"
    },
    {
        "20250528-探野T4-3D打印手板加工清单": 14,
        "Unnamed_Column_2": "镜头转接件1",
        "Unnamed_Column_3": "null",
        "Unnamed_Column_4": "镜头转接件1",
        "Unnamed_Column_5": "树脂",
        "Unnamed_Column_6": 1,
        "Unnamed_Column_7": "无",
        "Unnamed_Column_8": "null",
        "image_file": "image_row_16.png"
    },
    {
        "20250528-探野T4-3D打印手板加工清单": 15,
        "Unnamed_Column_2": "镜头转接件2",
        "Unnamed_Column_3": "null",
        "Unnamed_Column_4": "镜头转接件2",
        "Unnamed_Column_5": "树脂",
        "Unnamed_Column_6": 1,
        "Unnamed_Column_7": "无",
        "Unnamed_Column_8": "null",
        "image_file": "image_row_17.png"
    },
    {
        "20250528-探野T4-3D打印手板加工清单": 16,
        "Unnamed_Column_2": "镜头转接件3",
        "Unnamed_Column_3": "null",
        "Unnamed_Column_4": "镜头转接件3",
        "Unnamed_Column_5": "树脂",
        "Unnamed_Column_6": 1,
        "Unnamed_Column_7": "无",
        "Unnamed_Column_8": "null",
        "image_file": "image_row_18.png"
    },
    {
        "20250528-探野T4-3D打印手板加工清单": 17,
        "Unnamed_Column_2": "机芯支架",
        "Unnamed_Column_3": "null",
        "Unnamed_Column_4": "机芯支架",
        "Unnamed_Column_5": "树脂",
        "Unnamed_Column_6": 1,
        "Unnamed_Column_7": "无",
        "Unnamed_Column_8": "null",
        "image_file": "image_row_19.png"
    },
    {
        "20250528-探野T4-3D打印手板加工清单": 18,
        "Unnamed_Column_2": "散热板",
        "Unnamed_Column_3": "null",
        "Unnamed_Column_4": "散热板",
        "Unnamed_Column_5": "树脂",
        "Unnamed_Column_6": 1,
        "Unnamed_Column_7": "无",
        "Unnamed_Column_8": "null",
        "image_file": "image_row_20.png"
    },
    {
        "20250528-探野T4-3D打印手板加工清单": 19,
        "Unnamed_Column_2": "手柄",
        "Unnamed_Column_3": "null",
        "Unnamed_Column_4": "手柄",
        "Unnamed_Column_5": "树脂",
        "Unnamed_Column_6": 1,
        "Unnamed_Column_7": "无",
        "Unnamed_Column_8": "null",
        "image_file": "image_row_21.png"
    },
    {
        "20250528-探野T4-3D打印手板加工清单": 20,
        "Unnamed_Column_2": "探测器板",
        "Unnamed_Column_3": "null",
        "Unnamed_Column_4": "探测器板",
        "Unnamed_Column_5": "树脂",
        "Unnamed_Column_6": 1,
        "Unnamed_Column_7": "无",
        "Unnamed_Column_8": "null",
        "image_file": "image_row_22.png"
    },
    {
        "20250528-探野T4-3D打印手板加工清单": 21,
        "Unnamed_Column_2": "显示屏",
        "Unnamed_Column_3": "null",
        "Unnamed_Column_4": "显示屏",
        "Unnamed_Column_5": "树脂",
        "Unnamed_Column_6": 1,
        "Unnamed_Column_7": "无",
        "Unnamed_Column_8": "null",
        "image_file": "image_row_23.png"
    },
    {
        "20250528-探野T4-3D打印手板加工清单": 22,
        "Unnamed_Column_2": "显示屏后壳",
        "Unnamed_Column_3": "null",
        "Unnamed_Column_4": "显示屏壳体",
        "Unnamed_Column_5": "树脂",
        "Unnamed_Column_6": 1,
        "Unnamed_Column_7": "无",
        "Unnamed_Column_8": "null",
        "image_file": "image_row_24.png"
    },
    {
        "20250528-探野T4-3D打印手板加工清单": 23,
        "Unnamed_Column_2": "右饰片",
        "Unnamed_Column_3": "null",
        "Unnamed_Column_4": "右饰片",
        "Unnamed_Column_5": "树脂",
        "Unnamed_Column_6": 1,
        "Unnamed_Column_7": "无",
        "Unnamed_Column_8": "null",
        "image_file": "image_row_25.png"
    },
    {
        "20250528-探野T4-3D打印手板加工清单": 24,
        "Unnamed_Column_2": "左饰片",
        "Unnamed_Column_3": "null",
        "Unnamed_Column_4": "左饰片",
        "Unnamed_Column_5": "树脂",
        "Unnamed_Column_6": 1,
        "Unnamed_Column_7": "无",
        "Unnamed_Column_8": "null",
        "image_file": "image_row_26.png"
    },
    {
        "20250528-探野T4-3D打印手板加工清单": 25,
        "Unnamed_Column_2": "装饰片1",
        "Unnamed_Column_3": "null",
        "Unnamed_Column_4": "装饰片1",
        "Unnamed_Column_5": "树脂",
        "Unnamed_Column_6": 1,
        "Unnamed_Column_7": "无",
        "Unnamed_Column_8": "null",
        "image_file": "image_row_27.png"
    },
    {
        "20250528-探野T4-3D打印手板加工清单": 26,
        "Unnamed_Column_2": "装饰片2",
        "Unnamed_Column_3": "null",
        "Unnamed_Column_4": "装饰片2",
        "Unnamed_Column_5": "树脂",
        "Unnamed_Column_6": 1,
        "Unnamed_Column_7": "无",
        "Unnamed_Column_8": "null",
        "image_file": "image_row_28.png"
    },
    {
        "20250528-探野T4-3D打印手板加工清单": 27,
        "Unnamed_Column_2": "转接座",
        "Unnamed_Column_3": "null",
        "Unnamed_Column_4": "转接座",
        "Unnamed_Column_5": "树脂",
        "Unnamed_Column_6": 1,
        "Unnamed_Column_7": "无",
        "Unnamed_Column_8": "null",
        "image_file": "image_row_29.png"
    },
    {
        "20250528-探野T4-3D打印手板加工清单": 28,
        "Unnamed_Column_2": "转轴座",
        "Unnamed_Column_3": "null",
        "Unnamed_Column_4": "转轴座",
        "Unnamed_Column_5": "树脂",
        "Unnamed_Column_6": 1,
        "Unnamed_Column_7": "无",
        "Unnamed_Column_8": "null",
        "image_file": "image_row_30.png"
    },
    {
        "20250528-探野T4-3D打印手板加工清单": 29,
        "Unnamed_Column_2": "主板",
        "Unnamed_Column_3": "null",
        "Unnamed_Column_4": "主板",
        "Unnamed_Column_5": "树脂",
        "Unnamed_Column_6": 1,
        "Unnamed_Column_7": "无",
        "Unnamed_Column_8": "null",
        "image_file": "image_row_31.png"
    },
    {
        "20250528-探野T4-3D打印手板加工清单": 30,
        "Unnamed_Column_2": "主壳体",
        "Unnamed_Column_3": "null",
        "Unnamed_Column_4": "主壳体",
        "Unnamed_Column_5": "树脂",
        "Unnamed_Column_6": 1,
        "Unnamed_Column_7": "无",
        "Unnamed_Column_8": "null",
        "image_file": "image_row_32.png"
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
- DON’T modify or lose the "image_file" reference. Clearly pass it along unchanged.
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

# print Gemini standardized response clearly
print(response.text)