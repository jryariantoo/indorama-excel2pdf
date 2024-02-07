from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import Paragraph
import pandas as pd
import os
import configparser

def read_config():
    config = configparser.ConfigParser()
    config.read('config.ini')
    return config

def fetchdata(company_logo, company_name, month, year, usd, filename, start_row, start_col):

    ###################### save file 97-2003 format (.xls) 
    df = pd.read_excel(filename, sheet_name='Report Payroll')

    # Fetch data starting from row 17, column B
    row = start_row   # Row 17 (0-indexed)
    col = start_col  # Column B (0-indexed)
    df = df.iloc[row:, col:]


    ## cleaning
    cleaned_df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)

    # Set the first row as the column names
    cleaned_df.columns = cleaned_df.iloc[0]

    # Drop the first row (it's now the column names)
    cleaned_df = cleaned_df.drop(cleaned_df.index[0])

    # Reset the index after dropping the first row
    cleaned_df = cleaned_df.reset_index(drop=True)


    company_data = {
        'name': company_name,
        'logo': company_logo,
        'month': f"Salary Slip of {month} {year}" ## pake configparser
    }

    # Initialize an empty list to store rows and take_home_pay
    rows_data = []

    # Iterate over the rows of cleaned_df
    for index, row in cleaned_df.iterrows():
        if index == len(cleaned_df.index) - 1:
            break

        left_data = {
            'Employee ID': row['ID karyawan'],
            'Name': row['Nama karyawan'],
            'Position': row['Nama jabatan'],
            'Workplace': row['Nama departemen']
        }

        right_data = {
            'PTKP Status': row['Status nikah PTKP']
        }

        income_data = {
            'Salary': "${:,.2f}".format(round(row['Salary'] / usd, 2)),
            'Fixed Allowance': "${:,.2f}".format(round(row['Fixed Allowance'] / usd, 2))
        }

        deduction_data = {
            'Income Tax': "${:,.2f}".format(round(row['Income Tax'] / usd, 2)),
            "Employee's Old Age Security": "${:,.2f}".format(round(row["Employee’s Old Age Security"] / usd, 2)),
            "Employee's Pension Security": "${:,.2f}".format(round(row["Employee’s Pension Security"] / usd, 2)),
            "Employee's Health Security": "${:,.2f}".format(round(row["Employee’s Health Security"] / usd, 2)),
            'DPLK Employee': "${:,.2f}".format(round(row["DPLK Employe"] / usd, 2))
        }

        payment_data = {
            'Amount transfer': None,
            'Bank name': row['Nama bank 1'],
            'Account number': row['Nomor rekening bank 1']
        }

        unreceived_income = {
            "Company's Work Accident Security": "${:,.2f}".format(round(row['Company’s Work Accident Security'] / usd, 2)),
            "Company's Death Security": "${:,.2f}".format(round(row['Company’s Death Security'] / usd, 2)),
            "Company's Old Age Security": "${:,.2f}".format(round(row['Company’s Old Age Security'] / usd, 2)),
            "Company's Pension Security": "${:,.2f}".format(round(row['Company’s Pension Security'] / usd, 2)),
            "Company's Health Security": "${:,.2f}".format(round(row['Company’s Health Security'] / usd, 2)),
            "Tax borne by the Company": "${:,.2f}".format(round(row['Tax borne by the Company'] / usd, 2)),
            'DPLK Company': "${:,.2f}".format(round(row['DPLK Company'] / usd, 2))
        }

        row_data = {
            'employee':{
                'left': left_data,
                'right': right_data
            },
            'income': income_data,
            'deduction': deduction_data,
            'payment': payment_data,
            'unreceived_income': unreceived_income,
            'company' : company_data
        }

        # Append the dictionary to the list of rows
        rows_data.append(row_data)
    return rows_data        

def generate_payslip():
    #config setup
    config = read_config()
    company_name = config.get('Company', 'name')
    company_logo = config.get('Company', 'logo')
    month = config.get('Company', 'month')
    year = config.get('Company', 'year')
    row = int(config.get('Data', 'row'))
    col = int(config.get('Data', 'col'))
    usd = float(config.get('Rate', 'USD'))
    filename = config.get('Data', 'filename')

    rows_data = fetchdata(company_logo, company_name, month, year, usd, filename, row, col )
    
    output_directory = 'payslips'
    os.makedirs(output_directory, exist_ok=True)
    
    for data in rows_data:
        filename = f"Salary Slip of {month} {year} - {data['employee']['left']['Name']}.pdf"
        filepath = os.path.join(output_directory, filename)
        c = canvas.Canvas(filepath, pagesize=letter)

        # Header
        header(c, data['company'], data['employee']['left'], data['employee']['right'])
        c.line(11, 630, 602, 630)  # Line below the "Workplace" section

        # Body
        body(c, data['income'], data['deduction'], data['payment'], data['unreceived_income'])
        c.line(11, 155, 602, 155)

        # Footer
        footer(c, data['employee'])

        c.save()

def header(c, company, left, right):
    # Header - Logo and Company Name
    c.drawInlineImage(company['logo'], 5, 723, width=64, height=64)
    c.setFont("Helvetica-Bold", 16)
    c.drawString(75, 745, company['name'])
    c.drawString(75, 725, company['month'])

    # Subheader - Left Section
    c.setFont("Helvetica", 12)
    left_start_y = 720
    for i, (label, value) in enumerate(left.items(), start=1):
        c.drawString(10, left_start_y - i * 20, f"{label}")
        c.drawString(100, left_start_y - i * 20, value)  # Adjust the x-coordinate for cleaner spacing

       # Subheader - Right Section
    right_start_y = 720
    for i, (label, value) in enumerate(right.items(), start=1):
        if label == 'address':
            c.drawString(300, right_start_y - i * 20, f"{label}:")
            address_paragraph_style = ParagraphStyle(
                name='AddressParagraphStyle',
                fontName='Helvetica',  # Custom font name
                fontSize=12,  # Custom font size
                leading=14,  # Custom line spacing
            )
            address_paragraph = Paragraph(value, address_paragraph_style)
            address_paragraph.wrapOn(c, 220, 200)  # Wrap text to fit within 200 points width
            address_paragraph.drawOn(c, 370, right_start_y - i * 20 - 12)  # Draw wrapped text
        else:
            c.drawString(300, right_start_y - i * 20, f"{label}")
            c.drawString(380, right_start_y - i * 20, value)  # Adjust the x-coordinate for cleaner spacing
   

    # # Subheader - Right Section
    # right_start_y = 710
    # for i, (label, value) in enumerate(employee_data["right"].items(), start=1):
    #     c.drawString(290, right_start_y - i * 20, f"{label}")
    #     c.drawString(370, right_start_y - i * 20, value)  # Adjust the x-coordinate for cleaner spacing


def body(c, income, deduction, payment, unreceived_income):
    # Body - Income and Deduction
    c.setFont("Helvetica-Bold", 14)
    c.drawString(10, 610, "Income")
    c.setFont("Helvetica", 12)  # Set font weight to lower for income section
    for i, (item_name, item_value) in enumerate(income.items(), start=1):
        c.setFont("Helvetica", 12)        
        c.drawString(25, 605 - i * 20, f"{item_name}")
        if isinstance(item_value, (int, float)):  # Check if item_value is numeric
            c.drawString(140, 605 - i * 20, f"{item_value}")
        else:
            c.setFont("Helvetica-Bold", 12)
            c.drawString(140, 605 - i * 20, f"{item_value}")

    c.setFont("Helvetica-Bold", 14)
    c.drawString(300, 610, "Deduction")
    c.setFont("Helvetica", 12)  # Set font weight to lower for deduction section
    for i, (item_name, item_value) in enumerate(deduction.items(), start=1):
        c.setFont("Helvetica", 12)
        c.drawString(300, 605 - i * 20, f"{item_name}")
        if isinstance(item_value, (int, float)):  # Check if item_value is numeric
            c.drawString(470, 605 - i * 20, f"{item_value}")
        else:
            c.setFont("Helvetica-Bold", 12)
            c.drawString(470, 605 - i * 20, f"{item_value}")
    # Line below income and deduction
    c.line(11, 500, 602, 500)

    # Total Income and Deduction
    total_income = sum(float(value.replace('$', '').replace(',', '')) for value in income.values())
    c.setFont("Helvetica-Bold", 14)
    c.drawString(10, 485, "Total Income")
    c.drawString(140, 485, "${:,.2f}".format(total_income))

    total_deduction = sum(float(value.replace('$', '').replace(',', '')) for value in deduction.values())
    c.setFont("Helvetica-Bold", 14)
    c.drawString(300, 485, "Total Deduction")
    c.drawString(470, 485, "${:,.2f}".format(total_deduction))  

    # Line below total income and deduction
    c.line(11, 480, 602, 480)

    # Total take home pay
    take_home_pay = total_income - total_deduction
    c.setFont("Helvetica-Bold", 14)
    c.drawString(10, 465, "Take home pay")
    c.drawString(140, 465, "${:,.2f}".format(take_home_pay))  # Adjust the position as needed

    # Line below total income and deduction
    c.line(11, 460, 602, 460)

    c.setFont("Helvetica-Bold", 14)
    c.drawString(10, 440, "Payment")
    c.setFont("Helvetica-Bold", 12)  # Set font weight to lower for deduction section

    payment["Amount transfer"] = "${:,.2f}".format(take_home_pay)
    start_y = 440
    for i, (item_name, item_value) in enumerate(payment.items(), start=1):
        c.setFont("Helvetica", 12)
        c.drawString(10, start_y - i * 20, f"{item_name}")

        if isinstance(item_value, (int, float)):  # Check if item_value is numeric
            c.setFont("Helvetica-Bold", 12)
            c.drawString(140, start_y - i * 20, f"{item_value}")
        else:
            c.setFont("Helvetica-Bold", 12)
            c.drawString(140, start_y - i * 20, f"{item_value}")
    # Line below payment
    c.line(10, 370, 602, 370)

    c.setFont("Helvetica-Bold", 14)
    c.drawString(10, 350, "Unreceived income")
    c.setFont("Helvetica-Bold", 12)  # Set font weight to lower for deduction section
    start_y = 350
    for i, (item_name, item_value) in enumerate(unreceived_income.items(), start=1):
        c.setFont("Helvetica", 12)
        c.drawString(10, start_y - i * 20, f"{item_name}")
        c.setFont("Helvetica-Bold", 12)
        c.drawString(280, start_y - i * 20, f"{item_value}")

    # Line below unreceived income
    c.line(280, 205, 400, 205)

    total_unreceived_income = sum(float(value.replace('$', '').replace(',', '')) for value in unreceived_income.values())
    c.setFont("Helvetica-Bold", 14)
    c.drawString(10, 185, "Total unreceived")
    c.drawString(280, 185, "${:,.2f}".format(total_unreceived_income))   # Adjust the position as needed

    c.line(11, 180, 602, 180)

    ## Total unreceived and received income

    total_income_received_unreceived = take_home_pay + total_unreceived_income
    c.setFont("Helvetica-Bold", 14)
    c.drawString(10, 160, "Total income (received & unreceived)")
    c.drawString(280, 160, "${:,.2f}".format(total_income_received_unreceived))  # Adjust the position as needed

def footer(c, employee):
    # Footer
    c.setFont("Helvetica", 12)
    c.drawString(10, 130, "Given by")
    c.drawString(230, 130, "Known by")
    c.drawString(440, 130, "Received by")

    receivedby_y = 40
    c.drawString(440, receivedby_y, f"{employee['left']['Name']}")
    c.drawString(440, receivedby_y-15, f"{employee['left']['Position']}")




# Generate the payslip
# generate_payslip(employee_data_right, employee_data_left, company, income_items, deduction_items, unreceived_income)
generate_payslip()