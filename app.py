# Insert your python code here
# IMPORT PACKAGES IN THIS CELL

import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows

# INSERT ANSWER FOR QUESTION 2 HERE

product_descriptions = {}
product_prices = {}

df_descriptions = pd.read_excel(directory+'/'+ 'ProductData.xlsx', sheet_name="Descriptions")
for index,row in df_descriptions.iterrows():
    description = row.Description
    unit_number = row['Unit number']
    product_descriptions[unit_number] = description

df_prices = pd.read_excel(directory+'/'+ 'ProductData.xlsx', sheet_name="Prices")
for index,row in df_prices.iterrows():
    price = row['Price per unit']
    unit_number = row['Unit number']
    product_prices[unit_number] = price

# display(product_prices)
# display(product_descriptions)

# INSERT ANSWER FOR QUESTION 3 HERE

def get_order_data(invoice_num,filename,directory):
    invoice = openpyxl.load_workbook(directory+filename)
    selected_sheet = invoice['OrderForm']
    order_data = {}
    order_data["InvoiceNum"] = str(invoice_num)
    date = selected_sheet['C9'].value
    formatted_date = date.strftime("%Y/%m/%d")
    order_data["Date"] = formatted_date
    price0 = product_prices[selected_sheet['B24'].value]
    quantity0 = selected_sheet['C24'].value
    amount0 = ((product_prices[selected_sheet['B24'].value])*selected_sheet['C24'].value)
    desc0 = product_descriptions[selected_sheet['B24'].value]
    try:
        price1 = product_prices[selected_sheet['B25'].value]
        quantity1 = selected_sheet['C25'].value
        amount1 = ((product_prices[selected_sheet['B25'].value])*selected_sheet['C25'].value)
        desc1 = product_descriptions[selected_sheet['B25'].value]
    except (KeyError):
        pass
    try:
        price2 = product_prices[selected_sheet['B26'].value]
        quantity2 = selected_sheet['C26'].value
        amount2 = ((product_prices[selected_sheet['B26'].value])*selected_sheet['C26'].value)
        desc2 = product_descriptions[selected_sheet['B26'].value]
    except (KeyError):
        pass
    try:
        price3 = product_prices[selected_sheet['B27'].value]
        quantity3 = selected_sheet['C27'].value
        amount3=((product_prices[selected_sheet['B27'].value])*selected_sheet['C27'].value)
        desc3 = product_descriptions[selected_sheet['B27'].value]
    except (KeyError):
        pass
    try:
        price4 = product_prices[selected_sheet['B28'].value]
        quantity4 = selected_sheet['C28'].value
        amount4=((product_prices[selected_sheet['B28'].value])*selected_sheet['C28'].value)
        desc4 = product_descriptions[selected_sheet['B28'].value]
    except:
        pass
    try:
        order_data["Total"] = amount0+amount1+amount2+amount3+amount4
    except:
        pass
    try:
        order_data["Total"] = amount0+amount1+amount2+amount3
    except:
        pass
    try:
        order_data["Total"] = amount0+amount1+amount2
    except:
        pass
    try:
        order_data["Total"] = amount0+amount1
    except:
        order_data["Total"] = amount0
    order_data["CustomerAddress1"] = selected_sheet['C16'].value
    order_data["CustomerName"] = selected_sheet['C14'].value
    customer_name = selected_sheet['C14'].value
    order_data["CustomerCompany"] = selected_sheet['C15'].value
    order_data["CustomerAddress2"] = str(selected_sheet['C17'].value) + ", " + str(selected_sheet['C18'].value)+ " " + str(selected_sheet['C19'].value)
    customer_email = selected_sheet['C13'].value
    order_data["Price0"] = price0
    order_data["Quantity0"] = quantity0
    order_data["Amount0"] = amount0
    try:
        order_data["Price1"] = price1
        order_data["Quantity1"] = quantity1
        order_data["Amount1"] = amount1
    except:
        order_data["Price1"] = ""
        order_data["Quantity1"] = ""
        order_data["Amount1"] = ""
    try:
        order_data["Price2"] = price2
        order_data["Quantity2"] = quantity2
        order_data["Amount2"] = amount2
    except:
        order_data["Price2"] = ""
        order_data["Quantity2"] = ""
        order_data["Amount2"] = ""
    try:
        order_data["Price3"] = price3
        order_data["Quantity3"] = quantity3
        order_data["Amount3"] = amount3
    except:
        order_data["Price3"] = ""
        order_data["Quantity3"] = ""
        order_data["Amount3"] = ""
    try:
        order_data["Price4"] = price4
        order_data["Quantity4"] = quantity4
        order_data["Amount4"] = amount4
    except:
        order_data["Price4"] = ""
        order_data["Quantity4"] = ""
        order_data["Amount4"] = ""
    order_data["Description0"]= desc0
    try: 
        order_data["Description3"]=desc3
    except:
        order_data["Description3"]=""
    try:
        order_data["Description2"]=desc2
    except:
        order_data["Description2"]=""
    try:
        order_data["Description1"]=desc1
    except:
        order_data["Description1"]=""
    try:
        order_data["Description4"]=desc4
    except:
        order_data["Description4"]=""
    for key, value in order_data.items():
        if isinstance(value, (int, float)):
            order_data[key] = f'${value:,.2f}'
    return order_data, customer_name, customer_email, formatted_date

# df = get_pdf('MasterInvoice.pdf')
# df

directory = "/Users/briannaday/Downloads/Advanced Analytics/17 - Project 3 - Robotic Process Automation/"
filename = "Order1.xlsx"

get_order_data(8515,"Order1.xlsx", "/Users/briannaday/Downloads/Advanced Analytics/17 - Project 3 - Robotic Process Automation/")

# INSERT ANSWER FOR QUESTION 4 HERE
order_data = get_order_data(8515,"Order1.xlsx", "/Users/briannaday/Downloads/Advanced Analytics/17 - Project 3 - Robotic Process Automation/")
def create_invoice(invoice_num, order_data, directory):
   create_pdf(directory+'MasterInvoice.pdf', f'Invoice {invoice_num}.pdf', order_data[0])

create_invoice(8515, order_data,directory)

# INSERT ANSWER FOR QUESTION 5 HERE

gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()

def send_email(invoice_num, name, email, date, directory):
    gmail_user = "briannajday@gmail.com"
    recipients = [f"{email}"]
    subject = f"Your Order from PythonTech Co. - Invoice {invoice_num}"
    name = name
    body = f"""Dear {name},
    
    Attached is your invoice dated {date}.

    Thank you for your business.

    PythonTech Co.
    10 Python Drive
    Pleasant Grove, UT 84062
    Phone: (801) 555-2372
    Email: sales@pythontechco.com"""
    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)
    path_to_file = os.getcwd()+'/Invoice Template.xlsx' # MAC
    message = add_attachment(directory+f'Invoice {invoice_num}.pdf', message)
    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)
    mail_server.send_message(message)
    mail_server.quit()

send_email(8515,"Thomas Holt", "briannajday@gmail.com",'11/13/2023',directory)

# INSERT ANSWER FOR QUESTION 6 HERE

filenames = sorted(os.listdir(directory))
for filename in filenames:
    if 'Order' in filename and 'xlsx' in filename:
        invoice_num = 8515
        order_data = get_order_data(invoice_num,filename,directory)
        create_invoice(invoice_num,order_data,directory)
        name = order_data[1]
        email = order_data[2]
        date = order_data [3]
        send_email(invoice_num,name,email,date,directory)
        invoice_num = 8515 + 1