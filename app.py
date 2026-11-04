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

# INSERT ANSWER FOR QUESTION 1 HERE
while True:
    directory = single_input("Please input the path to your directory:")
    if directory[-1] != '\\':
        directory += '\\'
    if (os.path.exists(directory + 'ProductData.xlsx') and
        os.path.exists(directory + 'MasterInvoice.pdf') and
        os.path.exists(directory + 'Order1.xlsx') and
        os.path.exists(directory + 'Order2.xlsx')):
        break
    else:
        show_message('Error', 'The path you selected is incorrect. Please try again.')

# INSERT ANSWER FOR QUESTION 2 HERE
product_descriptions = {}
product_price = {}

#I need to add the reference to directory once I get it working.
descriptions = pd.read_excel('ProductData.xlsx',sheet_name = 'Descriptions')
prices = pd.read_excel('ProductData.xlsx',sheet_name = 'Prices')

for index, row in descriptions.iterrows():
    product_descriptions[row['Unit number']] = row['Description']

for index, row in prices.iterrows():
    product_price[row['Unit number']] = row['Price per unit']

df = get_pdf('MasterInvoice.pdf')
df

def get_order_data(invoice_num,filename,directory): #need to add directory once that's working
    order_data_dict = {}
    template = openpyxl.load_workbook(directory + filename)
    selected_sheet = template['OrderForm']
    name = selected_sheet['C14'].value
    email = selected_sheet['C13'].value
    date = selected_sheet['C9'].value.strftime('%m/%d/%Y')
    product0 = selected_sheet['B24'].value
    product1 = selected_sheet['B25'].value
    product2 = selected_sheet['B26'].value
    if product0 != None:
        price0 = product_price[product0]
    else:
        price0 = 0
    if product1 != None:
        price1 = product_price[product1]
    else:
        price1 = 0
    if product2 != None:
        price2 = product_price[product2]
    else:
        price2 = 0
    if product0 != None:
        quantity0 = selected_sheet['C24'].value
    else:
        quantity0 = 0
    if product1 != None:
        quantity1 = selected_sheet['C25'].value
    else:
        quantity1 = 0
    if product2 != None:
        quantity2 = selected_sheet['C26'].value
    else:
        quantity2 = 0
    order_data_dict['InvoiceNum'] = invoice_num
    order_data_dict['Date'] = selected_sheet['C9'].value.strftime('%m/%d/%Y')
    order_data_dict['Total'] = (price0 * quantity0) + (price1 * quantity1) + (price2 * quantity2)
    order_data_dict['CustomerAddress1'] = f'{str(selected_sheet['C16'].value)}'
    order_data_dict['CustomerAddress2'] = f'{str(selected_sheet['C17'].value)}, {str(selected_sheet['C18'].value)} {str(selected_sheet['C19'].value)}'
    order_data_dict['CustomerName'] = selected_sheet['C14'].value
    order_data_dict['CustomerCompany'] = selected_sheet['C15'].value
    order_data_dict['Price0'] = str(price0)
    order_data_dict['Quantity0'] = str(quantity0)
    amount0 = round(quantity0*price0,2)
    amount0 = f'{amount0:.2f}'
    amount1 = round(quantity1*price1,2)
    amount1 = f'{amount1:.2f}'
    amount2 = round(quantity2*price2,2)
    amount2 = f'{amount2:.2f}'
    if amount0 != '0.00':
        order_data_dict['Amount0']= str(amount0)
    else:
        order_data_dict['Amount0'] = ''
        order_data_dict['Price0'] = ''
        order_data_dict['Quantity0'] = ''
    order_data_dict['Price1'] = str(price1)
    order_data_dict['Quantity1'] = str(quantity1)
    if amount1 != '0.00':
        order_data_dict['Amount1']= str(amount1)
    else:
        order_data_dict['Amount1'] = ''
        order_data_dict['Price1'] = ''
        order_data_dict['Quantity1'] = ''
    order_data_dict['Price2'] = str(price2)
    order_data_dict['Quantity2'] = str(quantity2)
    if amount2 != '0.00':
        order_data_dict['Amount2']= str(amount2)
    else:
        order_data_dict['Amount2'] = ''
        order_data_dict['Price2'] = ''
        order_data_dict['Quantity2'] = ''
    try:
        order_data_dict['Description0'] = (product_descriptions[product0])
    except: 
        order_data_dict['Description0'] = ''
    try:
        order_data_dict['Description1'] = (product_descriptions[product1])
    except: 
        order_data_dict['Description1'] = '' 
    try: 
        order_data_dict['Description2'] = (product_descriptions[product2])
    except: 
        order_data_dict['Description2'] = ''
    return order_data_dict, name, email, date

order_data_dict, name, email, date= get_order_data(8515,'Order1.xlsx', directory)

# INSERT ANSWER FOR QUESTION 4 HERE
def create_invoice(invoice_num,order_data, directory): 
    create_pdf(directory + 'MasterInvoice.pdf', directory + f'Invoice {invoice_num}.pdf', order_data)

create_invoice(8515, order_data_dict, directory)

# INSERT ANSWER FOR QUESTION 5 HERE
def send_email(invoice_num,name,email,date,directory):
    # Enter your gmail user name and password
    
    gmail_user = "21loyler@gmail.com"
    gmail_password_path = select_file("Please select the gmail password file")
    gmail_password = open(gmail_password_path).read()

    # List of recipients

    recipients = [f'{email}']

    # Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    # Body of the email

    body = f"""Dear {name},
        
Attached is your invoice dated {date}.

Thank you for your business.
        
Python Tech Co.
10 Python Drive
Pleasant Grove, Ut 84062
Phone: (801) 555-2372
Email: sales@pythontechco.com"""

    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    path_to_file = f'{directory}Invoice {invoice_num}.pdf'

    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()

send_email(8515,'Thomas Holt','21loyler@gmail.com','11/13/2023',directory)

# INSERT ANSWER FOR QUESTION 6 HERE
filenames = os.listdir(directory)
filenames

invoice_num = 8515
for file in filenames:
    if "Order" in file and "xlsx" in file:
        order_data_dict, name, email, date= get_order_data(invoice_num,file,directory)
        create_invoice(invoice_num,order_data_dict,directory)
        send_email(invoice_num, name,email,date,directory)
        invoice_num += 1