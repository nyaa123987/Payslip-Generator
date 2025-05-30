import os
import pandas as pd
from fpdf import FPDF
import yagmail
from dotenv import load_dotenv

# Load environment variables
load_dotenv()
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

# Create a folder for payslips if it doesn't exist
if not os.path.exists("payslips"):
    os.makedirs("payslips")

# Read the Excel file
try:
    df = pd.read_excel(r"C:\Users\uncommonstudent\Desktop\python\employees.xlsx.xlsx")
except Exception as e:
    print("Error reading Excel file:", e)
    exit()

# Function to create a table
def create_table(pdf, headers, data):
    pdf.set_font("Times", 'B', 12)
    # Create header
    for header in headers:
        pdf.cell(40, 10, header, border=1, align='C')
    pdf.ln()

    pdf.set_font("Times", '', 12)
    # Create rows
    for row in data:
        for item in row:
            pdf.cell(40, 10, str(item), border=1, align='C')
        pdf.ln()

# Generate payslips and calculate net salary
for index, row in df.iterrows():
    try:
        emp_id = row['Employee ID']
        name = row['Name']
        email = row['Email']
        basic = row['Basic Salary']
        allowances = row['Allowances']
        deductions = row['Deductions']

        net_salary = basic + allowances - deductions

        pdf = FPDF()
        pdf.add_page()
        
       # Header
        pdf.set_fill_color(0, 102, 204)  # Blue color
        pdf.set_font("Times", 'B', 20)
        pdf.set_text_color(0, 102, 204)  # White text
        pdf.cell(200, 10, f"Payslip for {name}", ln=True, align='C')

         # Add second heading
        pdf.set_font("Times", 'B', 15)  # Font size for the second heading
        pdf.cell(0, 10, "Nyaa Technologies", ln=True, align='C')  # Centered second heading
        
        # Set text color to black for tables
        pdf.set_text_color(0, 0, 0)  # Black color for text

        # First Table: Employee Details
        pdf.set_font("Times", 'B', 14)
        pdf.cell(0, 10, "Employee Details", ln=True, align='L')
        pdf.set_font("Times", '', 12)
        create_table(pdf, ["Employee ID", "Name"], [[emp_id, name]])
        pdf.ln(5)

        # Second Table: Salary Details
        pdf.set_font("Times", 'B', 14)
        pdf.cell(0, 10, "Salary Details", ln=True, align='L')
        pdf.set_font("Times", '', 12)
        create_table(pdf, ["Basic Salary", "Allowances", "Deductions", "Net Salary"], [[basic, allowances, deductions, net_salary]])

        # Set footer color to blue
        pdf.set_text_color(0, 102, 204)  # Blue color for footer
        pdf.set_font("Times", '', 10)
        pdf.ln(10)
        pdf.cell(0, 10, "Thank you for your hard work!", ln=True, align='C')

        pdf_path = f"payslips/{emp_id}.pdf"
        pdf.output(pdf_path)
        print(f"Payslip generated for {name}")

    except Exception as e:
        print(f"Error generating payslip for {row['Name']}: {e}")

# Send emails
try:
    yag = yagmail.SMTP(EMAIL_USER, EMAIL_PASSWORD)
    for index, row in df.iterrows():
        try:
            emp_id = row['Employee ID']
            name = row['Name']
            email = row['Email']
            pdf_path = f"payslips/{emp_id}.pdf"

            yag.send(
                to=email,
                subject="Your Payslip for This Month",
                contents=f"Hello {name},\n\nPlease find your payslip attached.\n\nBest regards,\nHR Team",
                attachments=pdf_path
            )
            print(f"Email sent to {name} at {email}")

        except Exception as e:
            print(f"Failed to send email to {row['Name']}: {e}")
except Exception as e:
    print("Failed to set up email client:", e)

