from dotenv import load_dotenv
import os
import openai
import pandas as pd
import PyPDF2
import re
import openpyxl
import os
load_dotenv()

# Set your OpenAI API key
apiKey = os.getenv("API_KEY")

if apiKey is None:
    raise ValueError(
        "API_KEY environment variable not found. Please set it in your .env file.")

openai.api_key = apiKey


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

# The prompt can be modified


def create_prompt(resume_text, job_description):
    prompt = f"""
You are an AI assistant that extracts specific information from resumes for HR purposes. You will extract every possible piece of
Information requested in this prompt. You will be meticulous and detail-oriented. I will provide the resume text and a job description.
In your summary, you will summarize each candidate's job history and experience and based on the provided job_description,
you will provide a rating of how the candidate fits or does not fit within the provided description. 


Given the following resume text:
"{resume_text}"

And the job description:
"{job_description}"

Extract the following information in the specified template:

- Name:
- Current Job Title: (ONLY IF A current job exists or a title is mentioned - if the person is a college student or recent graduate then N/A)
- Current Company (if not available then N/A. For students who are are currently in university, it must be N/A. For other non-students, retrieve the most recent job title regardless):
- Phone Number: (N/A if it does not exist)
- Email: 
- Link to profile on LinkedIn:
- Education: (For education, please put the university that is or has been attended by the candidate. Please)
- 2-3 sentence summary of their education and career:
- Rating: (based on the provided job description)

Provide the output in the exact same format as above.
"""
    return prompt


def get_candidate_info(prompt):
    response = openai.ChatCompletion.create(
        model="gpt-4o",
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
        'Email': '',
        'LinkedIn Profile': '',
        'Education': '',
        'Summary': '',
        'Rating': ''
    }

    # Use regular expressions to extract each piece of information
    patterns = {
        'Name': r'- Name:\s*(.*)',
        'Current Job Title': r'- Current Job Title:\s*(.*)',
        'Current Company': r'- Current Company:\s*(.*)',
        'Phone Number': r'- Phone Number:\s*(.*)',
        'Email': r'- Email: \s*(.*)',
        'LinkedIn Profile': r'- Link to profile on LinkedIn:\s*(.*)',
        'Education': r'- Education:\s*(.*)',
        'Summary': r'- 2-3 sentence summary of their education and career:\s*(.*)',
        'Rating': r'- Rating:\s*(.*)'
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
    job_description = """Responsibilities ▪ Financial Leadership & Strategic Planning: Develop and execute
    the financial strategy for both the Group and NCA, aligning
    financial and business goals to drive growth, enhance
    profitability, and ensure long-term sustainability.
    ▪ Investment & Growth Opportunities: Identify and evaluate
    investment opportunities and partnerships that align with the
    Group's and NCA’s strategic objectives, overseeing due diligence
    and integration processes.
    ▪ Risk Management & Compliance: Develop and oversee risk
    management policies and compliance frameworks to mitigate
    financial and operational risks, ensuring adherence to legal and
    regulatory requirements.
    ▪ Strategic Oversight of Financial Operations: Provide strategic
    oversight of all financial functions including financing, treasury,
    financial planning and analysis, and accounting to ensure
    efficiency, accuracy, and transparency.
    ▪ Technology & Innovation in Finance: Champion the adoption of
    technological innovations and best practices in financial
    management and IT systems to improve efficiency, data analytics,
    and business intelligence.
    Operational Responsibilities:
    ▪ Corporate Governance & Internal Controls: Reinforce corporate
    governance and internal control systems to support strategic
    decisions and protect organizational assets.
    ▪ Stakeholder Engagement & Communication: Maintain strong
    relationships with all stakeholders, promoting effective
    communication and strategic alignment.
    CONFIDENTIAL WORK DOCUMENT - FOR INTERNAL USE ONLY 5
    Chief Financial Officer
    ▪ Leadership & Team Development: Cultivate and mentor a high-
    performing finance and IT team, encouraging a culture of
    excellence, innovation, and strategic initiative.
    Pro-active person who can take ownership of the position as Group CFO and
    CFO of NCA and be a trusted voice internally and externally and feel
    comfortable presenting facts and figures towards investors and
    stakeholders.
    Leadership Tasks As the GROUP CFO & NCA CFO, you will play a critical role in shaping the
    future of our company, leveraging your strategic insight and financial
    expertise to drive growth, innovation, and excellence. We are looking for a
    visionary leader passionate about making a significant impact.
    ▪ Manage and motivate the team.
    ▪ Ensure punctuality in reporting and comply with deadlines
    ▪ Act as Business Partner towards the organization
    ▪ Be trusted “wingman” to the CEO (Anders Peter)
    ▪ Provide leadership in financial management
    ▪ Strategic investment
    ▪ IT
    ▪ Operational excellence, with expertise in financial planning
    ▪ Risk management
    ▪ Corporate governance
    Key Tasks ▪ Implement effective reporting systems
    ▪ Consolidation
    ▪ 12-months cash-flow analysis
    ▪ Budgeting and forecasting
    ▪ Internal analysis
    ▪ Handle suppliers and customers from a finance perspective
    ▪ Conclude annual reports for subsidiaries within the Group and
    reconcile these with local auditors
    ▪ ERP upgrade
    ▪ Power BI
    ▪ “Growth Financing” – Each unit must be self-financing
    Tasks by percentage on a
    monthly basis
    Tasks by Percentage on a Monthly Basis
    •⁠  ⁠Financial Reporting and Analysis (25%): Prepare and review monthly
    financial statements, budgets, and forecasts to ensure accuracy and
    compliance with financial regulations.
    •⁠  ⁠Strategic Planning and Execution (20%): Develop and implement
    financial strategies aligned with company goals, including investment
    opportunities and risk management.
    CONFIDENTIAL WORK DOCUMENT - FOR INTERNAL USE ONLY 6
    Chief Financial Officer
    •⁠  ⁠Cashflow and Growth Financing (20%): Monitor and manage cash flow
    to ensure liquidity, identify and secure financing opportunities to
    support company expansion and growth initiatives.
    •⁠  ⁠Team Leadership and Development (15%): Mentor and oversee the
    finance team, ensuring high performance and continuous
    improvement.
    •⁠  ⁠Stakeholder Engagement (10%): Maintain strong relationships with
    banks, auditors, investors, and other key stakeholders, ensuring clear
    and effective communication.
    •⁠  ⁠Operational Oversight (10%): Monitor and enhance internal controls,
    IT systems, and financial processes to support operational excellence.
    Critical success factors (KPI) ▪ Improvement of reporting processes
    ▪ Delivering accurate reports on time
    ▪ Monthly reports on day xx at the moment
    ▪ Cash-flow management
    ▪ Month close day 10 as of today – must come down to day 5"""

    # Initialize a list to hold all candidate data
    candidates_data = []

    # Iterate over each PDF in the resumes folder
    for filename in os.listdir(resumes_folder):
        if filename.endswith('.pdf'):
            pdf_path = os.path.join(resumes_folder, filename)
            resume_text = extract_text_from_pdf(pdf_path)
            resume_text = preprocess_text(resume_text)
            print("RESUME TEXT: ", resume_text)
            prompt = create_prompt(resume_text, job_description)
            info_text = get_candidate_info(prompt)
            candidate_info = parse_candidate_info(info_text)
            candidates_data.append(candidate_info)
            print(f"Extracted information for {filename}")
            print(candidate_info)
    print(candidates_data)
    # Create a DataFrame and save to Excel
    df = pd.DataFrame(candidates_data)
    df.to_excel('c_candidates_info.xlsx', index=False)
    print("Candidates information has been successfully extracted and saved to 'candidates_info.xlsx'.")


if __name__ == '__main__':
    main()
