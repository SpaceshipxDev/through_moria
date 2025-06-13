
-? 

pump 

ingest.py -> think.py -? parse.py 


curl https://ark.cn-beijing.volces.com/api/v3/chat/completions \
  -H "Content-Type: application/json" \
  -H "Authorization: Bearer 37946450-91ae-4a14-a7e6-3728cc785904" \
  -d '{
    "model": "doubao-1.5-pro-32k-250115",
    "messages": [
        {
            "role": "system",
            "content": "You are a helpful assistant."
        },
        {
            "role": "user",
            "content": "Hello!"
        }
    ]
  }'