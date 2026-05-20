import os
from dotenv import load_dotenv
from google import genai

load_dotenv()

# Client auto-reads GEMINI_API_KEY from environment
# Falls back to prompting if not set
api_key = os.getenv("GEMINI_API_KEY")
if not api_key:
    api_key = input("Enter your Gemini API key: ").strip()

client = genai.Client(api_key=api_key)

print("Gemini test chat — type 'exit' to quit\n")

while True:
    prompt = input("You: ").strip()
    if prompt.lower() == "exit":
        break
    if not prompt:
        continue
    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt
    )
    print(f"Gemini: {response.text}\n")
