from dotenv import load_dotenv
from openai import OpenAI
import pandas as pd
import json
import pdfplumber
import os


load_dotenv()
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))







# Step-1 Extract pdf data using pdf library

def extract_text_from_pdf(pdf_path):
    full_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text += text + "\n"
    return full_text


if __name__ == "__main__":
    path = "Data/Input/sample.pdf"
    extracted_text = extract_text_from_pdf(path)

    
# Step-2 & 3 - Send extracted text to Open AI and get structured JSON response

response = client.chat.completions.create(
    model="gpt-4o-mini",
    messages=[
        {
            "role": "system",
            "content": "You are a helpful assistant that extracts information from resumes. Always respond in valid JSON format only, no extra text, no markdown, no backticks."
        },
        {
            "role": "user",
            "content": f"Extract the following details from this resume and return as JSON:\n- name\n- email\n- phone\n- skills\n- work_experience (list with company, role, duration, responsibilities)\n\nResume:\n{extracted_text}"
        }
    ]
)

raw_result = response.choices[0].message.content

# Step 3 - Parse the response as JSON

structured_data = json.loads(raw_result)
print(json.dumps(structured_data, indent=4))


# Step 4 - Save to Excel

output_path = "Data/Output/resume_output.xlsx"

# Sheet 1 - Basic Info
basic_info = {
    "Field": ["Name", "Email", "Phone", "Skills"],
    "Value": [
        structured_data["name"],
        structured_data["email"],
        structured_data["phone"],
        ", ".join(structured_data["skills"])
    ]
}
df_basic = pd.DataFrame(basic_info)

# Sheet 2 - Work Experience
work_exp = []
for job in structured_data["work_experience"]:
    work_exp.append({
        "Company": job["company"],
        "Role": job["role"],
        "Duration": job["duration"],
        "Responsibilities": "\n".join(job["responsibilities"])
    })
df_work = pd.DataFrame(work_exp)

# Save both sheets to Excel
with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    df_basic.to_excel(writer, sheet_name="Basic Info", index=False)
    df_work.to_excel(writer, sheet_name="Work Experience", index=False)

print(f"Excel file saved to {output_path}")