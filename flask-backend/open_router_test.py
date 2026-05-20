import os
from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()

api_key = (os.getenv("OPEN_ROUTER_API_KEY") or "").strip()
if not api_key:
    api_key = input("Enter your OpenRouter API key: ").strip()

print(f"Using key: {api_key[:12]}...{api_key[-4:]}  (len={len(api_key)})")

client = OpenAI(
    api_key=api_key,
    base_url="https://openrouter.ai/api/v1",
)

print("OpenRouter test chat (gpt-5-nano) — type 'exit' to quit\n")

while True:
    prompt = input("You: ").strip()
    if prompt.lower() == "exit":
        break
    if not prompt:
        continue
    response = client.chat.completions.create(
        model="openai/gpt-5-nano",
        messages=[{"role": "user", "content": prompt}],
    )
    print(f"GPT: {response.choices[0].message.content}\n")
