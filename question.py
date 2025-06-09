import os
from openai import OpenAI

client = OpenAI(
    # 从环境变量中读取您的方舟API Key
    api_key=os.environ.get("37946450-91ae-4a14-a7e6-3728cc785904"), 
    base_url="https://ark.cn-beijing.volces.com/api/v3",
    # 深度思考模型耗费时间会较长，建议您设置一个较长的超时时间，推荐为30分钟
    timeout=1800,
    )
response = client.chat.completions.create(
    # 替换 <Model> 为 Model ID
    model="Doubao-1.5-lite-32k",
    messages=[
        {"role": "user", "content": "我要研究深度思考模型与非深度思考模型区别的课题，怎么体现我的专业性"}
    ]
)
# 当触发深度思考时，打印思维链内容
if hasattr(response.choices[0].message, 'reasoning_content'):
    print(response.choices[0].message.reasoning_content)
print(response.choices[0].message.content)