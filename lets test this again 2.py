import os
import xlrd
import time
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

current_dir = os.path.dirname(os.path.abspath(__file__))

# Specify the file name and path relative to the current directory
filename = "clients.xls"
path = os.path.join(current_dir, filename)

openFile = xlrd.open_workbook(path)
sheet = openFile.sheet_by_index(0)

mail_list = []
amount = []
name = []

for k in range(1, sheet.nrows):
    client = sheet.cell_value(k, 0)
    email = sheet.cell_value(k, 1)
    paid = sheet.cell_value(k, 2)
    count_amount = sheet.cell_value(k, 3)
    if paid.lower() == 'no':
        mail_list.append(email)
        amount.append(count_amount)
        name.append(client)

email = 'greatmonunath99@gmail.com'
password = 'ndidtrawnvineutj'

server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(email, password)

try:
    for i, mail_to in enumerate(mail_list):
        clientName = name[i]
        amount_owed = f'₹{int(amount[i])}'

        subject = f'{clientName}, you have a new email'
        message = f'Dear {clientName},\n\n' \
                  f'We would like to inform you that you owe {amount_owed}.\n\n' \
                  'Best Regards'

        msg = MIMEMultipart()
        msg['From'] = email
        msg['To'] = mail_to
        msg['Subject'] = subject
        msg.attach(MIMEText(message, 'plain'))

        print(f"Sending email to: {mail_to}")
        server.sendmail(email, mail_to, msg.as_string())

    print('All emails have been sent successfully!')
except Exception as e:
    print(f"An error occurred while sending the emails: {str(e)}")

server.quit()
print('Process is finished!')
time.sleep(10)
