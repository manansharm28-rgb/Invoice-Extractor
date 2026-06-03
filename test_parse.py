import os
import json
import requests
import ssl

# Disable SSL verification globally
ssl._create_default_https_context = ssl._create_unverified_context
os.environ["PYTHONHTTPSVERIFY"] = "0"
os.environ["CURL_CA_BUNDLE"] = ""
os.environ["REQUESTS_CA_BUNDLE"] = ""

import httpx

from dotenv import load_dotenv
from llama_parse import LlamaParse

# Load environment variables
load_dotenv()

# Get keys
llama_key = os.environ.get("LLAMA_CLOUD_API_KEY")
groq_key = os.environ.get("GROQ_API_KEY")
pdf_path = r"c:\Users\A554471\Desktop\Invoice\INV NO,329873.pdf"

print("--- Step 1: Parsing PDF with LlamaParse ---")
try:
    instruction = (
        "This is an invoice document. Please parse it and extract ONLY the following details: "
        "1. Invoice Number (invoice_no)\n"
        "2. Invoice Date (invoice_date)\n"
        "3. Part Numbers (part_no) and their corresponding Part Quantities (part_quantity).\n"
        "Ignore all other details. Focus only on these specific columns/values and format them clearly as tables or lists."
    )
    
    parser = LlamaParse(
        api_key=llama_key,
        result_type="markdown",
        verbose=True,
        parsing_instruction=instruction
    )
    
    documents = parser.load_data(pdf_path)
    markdown_text = "\n\n".join([doc.text for doc in documents])
    
    print("\n--- Parsed Markdown Text (First 1000 characters) ---")
    print(markdown_text[:1000])
    
    # Save the markdown text
    with open("parsed_invoice.md", "w", encoding="utf-8") as f:
        f.write(markdown_text)
    print("\nSaved full markdown to parsed_invoice.md")
    
except Exception as e:
    print(f"LlamaParse error: {e}")
    exit(1)

print("\n--- Step 2: Extracting Structured Details using Groq ---")
try:
    url = "https://api.groq.com/openai/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {groq_key}",
        "Content-Type": "application/json"
    }
    
    prompt = f"""
    You are an expert data extraction agent. You will be provided with the markdown content of a parsed invoice document.
    Extract the following details:
    1. Invoice Number (invoice_no)
    2. Invoice Date (invoice_date)
    3. Line items, where each line item contains:
       - Part Number (part_no)
       - Quantity (part_quantity)
       
    Invoice Markdown Content:
    ---
    {markdown_text}
    ---
    
    Return ONLY a JSON object with this exact structure:
    {{
      "invoice_no": "str or null",
      "invoice_date": "str or null",
      "line_items": [
        {{
          "part_no": "str or null",
          "part_quantity": int or null
        }}
      ]
    }}
    
    Rules:
    - If a field is not found, set its value to null.
    - Extract multiple line items if present.
    - Do not add any explanation or conversational text. Return only valid JSON.
    """
    
    payload = {
        "model": "llama-3.3-70b-versatile",
        "messages": [
            {"role": "system", "content": "You are a precise data extraction agent that outputs JSON only."},
            {"role": "user", "content": prompt}
        ],
        "response_format": {"type": "json_object"},
        "temperature": 0
    }
    
    res = requests.post(url, headers=headers, json=payload, verify=False)
    res.raise_for_status()
    response_data = res.json()
    result_text = response_data['choices'][0]['message']['content']
    
    print("\n--- Groq Response JSON ---")
    print(result_text)
    
except Exception as e:
    print(f"Groq extraction error: {e}")
