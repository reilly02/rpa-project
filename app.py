# Insert your python code here
import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows|

# INSERT ANSWER FOR QUESTION 1 HERE

while True:
        directory = single_input("Please enter the directory where the files are located")
        if directory[-1] != '/':
            directory = directory + '/'
        filepath = os.listdir(directory)
        if 'ProductData.xlsx' in filepath and 'MasterInvoice.pdf' in filepath and 'Order1.xlsx' in filepath and 'Order2.xlsx' in filepath and 'Order3.xlsx' in filepath and 'Order4.xlsx' in filepath:
            break
        else:
            show_message("Error", "The directory you entered is incorrect. Please try again.")



# INSERT ANSWER FOR QUESTION 2 HERE
product_descriptions = {}
product_prices = {}

prices = pd.read_excel(directory+'ProductData.xlsx',sheet_name='Prices')
descriptions = pd.read_excel(directory+'ProductData.xlsx',sheet_name='Descriptions')

for d,description in descriptions.iterrows():
    product_descriptions[description['Unit number']] = description['Description']

for p,price in prices.iterrows():
    product_prices[price['Unit number']] = price['Price per unit']



# INSERT ANSWER FOR QUESTION 3 HERE
def get_order_data(invoice_num,filename,directory):
    template = openpyxl.load_workbook(directory + filename)
    selected_sheet = template['OrderForm']
    update_dict = {}
    update_dict['InvoiceNum'] = str(invoice_num)
    update_dict['Date'] = selected_sheet['C9'].value.strftime('%m/%d/%Y')
    update_dict['CustomerAddress1'] = selected_sheet['C16'].value
    update_dict['CustomerName'] = selected_sheet['C14'].value
    update_dict['CustomerCompany'] = selected_sheet['C15'].value
    update_dict['CustomerAddress2'] = f"{selected_sheet['C17'].value}, {selected_sheet['C18'].value} {selected_sheet['C19'].value}"
    name = selected_sheet['C14'].value
    email = selected_sheet['C13'].value
    date = selected_sheet['C9'].value.strftime('%m/%d/%Y')

    total = 0
    
    if selected_sheet['B24'].value is not None:
        amount0 = (selected_sheet['C24'].value*product_prices[selected_sheet['B24'].value])
        price0 = product_prices[selected_sheet['B24'].value]
        update_dict['Price0'] = f'${price0:,.2f}'
        update_dict['Quantity0'] = selected_sheet['C24'].value
        update_dict['Amount0'] = f'${amount0:,.2f}'
        update_dict['Description0'] = product_descriptions[selected_sheet['B24'].value]
        total = total + amount0
    if selected_sheet['B25'].value is not None:
        amount1 = (selected_sheet['C25'].value * product_prices[selected_sheet['B25'].value])
        price1 = product_prices[selected_sheet['B25'].value]
        update_dict['Price1'] = f'${price1:,.2f}'
        update_dict['Quantity1'] = selected_sheet['C25'].value
        update_dict['Amount1'] = f'${amount1:,.2f}'
        update_dict['Description1'] = product_descriptions[selected_sheet['B25'].value]
        total = total + amount1
    if selected_sheet['B26'].value is not None:
        amount2 = (selected_sheet['C26'].value * product_prices[selected_sheet['B26'].value])
        price2 = product_prices[selected_sheet['B26'].value]
        update_dict['Price2'] = f'${price2:,.2f}'
        update_dict['Quantity2'] = selected_sheet['C26'].value
        update_dict['Amount2'] = f'${amount2:,.2f}'
        update_dict['Description2'] = product_descriptions[selected_sheet['B26'].value]
        total = total + amount2
    if selected_sheet['B27'].value is not None:
        amount3 = (selected_sheet['C27'].value * product_prices[selected_sheet['B27'].value])
        price3 = product_prices[selected_sheet['B27'].value]
        update_dict['Price3'] = f'${price3:,.2f}'
        update_dict['Quantity3'] = selected_sheet['C27'].value
        update_dict['Amount3'] = f'${amount3:,.2f}'
        update_dict['Description3'] = product_descriptions[selected_sheet['B27'].value]
        total = total + amount3
    if selected_sheet['B28'].value is not None:
        amount4 = (selected_sheet['C28'].value * product_prices[selected_sheet['B28'].value])
        price4 = product_prices[selected_sheet['B28'].value]
        update_dict['Price4'] = f'${price4:,.2f}'
        update_dict['Quantity4'] = selected_sheet['C28'].value
        update_dict['Amount4'] = f'${amount4:,.2f}'
        update_dict['Description4'] = product_descriptions[selected_sheet['B28'].value]
        total = total + amount4

    update_dict['Total'] = f'${total:,.2f}'
    
    return update_dict, name, email, date



# INSERT ANSWER FOR QUESTION 4 HERE

def create_invoice(invoice_num, order_data, directory):

    template = directory + 'MasterInvoice.pdf'
    output = directory + f'Invoice {invoice_num}.pdf'
    
    create_pdf(template, output, order_data)



# INSERT ANSWER FOR QUESTION 5 HERE
def send_email(invoice_num, name, email, date, directory):

    gmail_user = "bacoccinireilly@gmail.com"
    gmail_password_path = select_file('Please select the gmail password file')
    gmail_password = open(gmail_password_path).read()


    # Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    # Body of the email


    body = f"""Dear {name},
        
    Attached is your invoice dated {date}.

    Thank you for your business.
        
    PythonTech Co.
    10 Python Drive
    Pleasant Grove, UT 84062
    Phone: (801) 555-2372
    Email: sales@pythontechco.com
    """

    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = email
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    path_to_file = directory + 'Invoice '+str(invoice_num)+  '.pdf'
    
    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()

    return


# INSERT ANSWER FOR QUESTION 6 HERE
filenames = os.listdir(directory)
filenames.sort()

invoice_num = 8515

for file in filenames:
    if 'Order' in file and 'xlsx' in file:
        order_data, name, email, date = get_order_data(invoice_num, file, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)
        invoice_num += 1
