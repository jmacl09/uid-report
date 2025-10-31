import os
from openai import AzureOpenAI

endpoint = "https://wan-syslog-agent-resource.cognitiveservices.azure.com/"
model_name = "gpt-5-mini"
deployment = "gpt-5-mini-2"

subscription_key = "EahXigmZo45i1mwPHeWGkvXEBtiTOJXIQAawEgKNMO8S3TGRDONZJQQJ99BJAC5RqLJXJ3w3AAAAACOG0hI5"
api_version = "2024-12-01-preview"

client = AzureOpenAI(
    api_version=api_version,
    azure_endpoint=endpoint,
    api_key=subscription_key,
)

response = client.chat.completions.create(
    messages=[
        {
            "role": "system",
            "content": "You are a helpful assistant.",
        },
        {
            "role": "user",
            "content": "I am going to Paris, what should I see?",
        }
    ],
    max_completion_tokens=16384,
    model=deployment
)

print(response.choices[0].message.content)
