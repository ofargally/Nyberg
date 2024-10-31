import os
import openai
import pandas as pd
import PyPDF2
import re
import openpyxl

# Set your OpenAI API key
openai.api_key = ''


def extract_text_from_pdf(pdf_path):
    text = ""
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        for page in reader.pages:
            text += page.extract_text()
        return text

def preprocess_text(text):
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

#The prompt can be modified 
def create_prompt(resume_text, job_description):
    prompt = f"""
You are an AI assistant that extracts specific information from resumes for HR purposes.

Given the following resume text:
"{resume_text}"

And the job description:
"{job_description}"

Extract the following information in the specified template:

- Name:
- Current Job Title: (ONLY IF A current job exists or a title is mentioned - if the person is a college student or recent graduate then N/A)
- Current Company (if not available then N/A):
- Phone Number:
- Link to profile on LinkedIn:
- Education (undergraduate and postgraduate college/university only):
- 2-3 sentence summary of their education and career:
- Rating (A-B-C) based on how relevant their experience is to the specific position provided.

Provide the output in the exact same format as above.
"""
    return prompt

def get_candidate_info(prompt):
    response = openai.ChatCompletion.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": "You are an AI assistant that extracts specific information from resumes for HR purposes."},
            {"role": "user", "content": prompt}
        ],
        max_tokens=700,
        temperature=0,
    )
    return response['choices'][0]['message']['content'].strip()

def parse_candidate_info(info_text):
    # Initialize a dictionary to hold the extracted info
    info = {
        'Name': '',
        'Current Job Title': '',
        'Current Company': '',
        'Phone Number': '',
        'LinkedIn Profile': '',
        'Education': '',
        'Summary': '',
        'Rating': ''
    }
    
    # Use regular expressions to extract each piece of information
    patterns = {
        'Name': r'- Name:\s*(.*)',
        'Current Job Title': r'- Current Job Title:\s*(.*)',
        'Current Company': r'- Current Company \(if not available then N/A\):\s*(.*)',
        'Phone Number': r'- Phone Number:\s*(.*)',
        'LinkedIn Profile': r'- Link to profile on LinkedIn:\s*(.*)',
        'Education': r'- Education \(undergraduate and postgraduate college/university only\):\s*(.*)',
        'Summary': r'- 2-3 sentence summary of their education and career:\s*(.*)',
        'Rating': r'- Rating \(A-B-C\) based on how relevant their experience is to the specific position provided\.\s*(.*)'
    }
    
    for key, pattern in patterns.items():
        match = re.search(pattern, info_text)
        if match:
            info[key] = match.group(1).strip()
    return info

def main():
    # Path to the folder containing resume PDFs
    resumes_folder = 'resumes/'
    # Job description for evaluating candidates
    job_description = "Your specific job description here."
    
    # Initialize a list to hold all candidate data
    candidates_data = []
    
    # Iterate over each PDF in the resumes folder
    for filename in os.listdir(resumes_folder):
        if filename.endswith('.pdf'):
            pdf_path = os.path.join(resumes_folder, filename)
            resume_text = extract_text_from_pdf(pdf_path)
            resume_text = preprocess_text(resume_text)
            prompt = create_prompt(resume_text, job_description)
            info_text = get_candidate_info(prompt)
            candidate_info = parse_candidate_info(info_text)
            candidates_data.append(candidate_info)
            print(f"Extracted information for {filename}")
            print(candidate_info)
    print(candidates_data)
    # Create a DataFrame and save to Excel
    df = pd.DataFrame(candidates_data)
    df.to_excel('candidates_info.xlsx', index=False)
    print("Candidates information has been successfully extracted and saved to 'candidates_info.xlsx'.")

if __name__ == '__main__':
    main()

