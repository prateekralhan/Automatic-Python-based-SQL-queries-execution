# imports for SQL data part
import pyodbc
from datetime import datetime, timedelta
import time
import pandas as pd
import os

# imports for sending email
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib
from plyer import notification
import xlsxwriter

out_dir = os.path.abspath('output_files/')

date = datetime.today() - timedelta(days=7)  # get the date 7 days ago

date = date.strftime("%Y-%m-%d")  # convert to format yyyy-mm-dd
timestr=time.strftime("%Y-%m-%d-%H-%M-%S")

# function to receive Windows Desktop notifications
def notifyme(title,message):
    notification.notify(
        title=title,
        message=message,
        app_icon = None,
        timeout=20
    )

cnxn_str = ("Driver={SQL Server Native Client 11.0};"
        "Server = < Enter the Server Address here > ;"
        "Database = < Enter the Database >;"
        "Trusted_Connection=yes;"
        "UID = < Enter the UserID here for accessing your db >;"
        "PWD= < Enter the password for the same. > ")

cnxn = pyodbc.connect(cnxn_str)  # initialise connection

# build up our query string
query = ("""SELECT * from Table_Name""")

# execute the query and read to a dataframe in Python
df = pd.read_sql(query, cnxn)

# saving the query results in a excel file in folder "output_files"
UT_XLS = "SQL_query_output_"+timestr+".xlsx"
writer = pd.ExcelWriter(os.path.join(out_dir,OUT_XLS), engine='xlsxwriter')
df.to_excel(writer,index=False,sheet_name='IRIS PROD Usage')

# formatting the excel spreadsheet a bit
workbook = writer.book
worksheet=writer.sheets['IRIS PROD Usage']

format=workbook.add_format({'text_wrap': True})

red = workbook.add_format({'bg_color':   '#FFC7CE',
                               'font_color': '#9C0006','text_wrap': True})
yellow = workbook.add_format({'bg_color':   '#FFEB9C',
                               'font_color': '#9C6500','text_wrap': True})
green = workbook.add_format({'bg_color':   '#C6EFCE',
                               'font_color': '#006100','text_wrap': True})
blue = workbook.add_format({'bg_color':   '#66CCFF',
                               'font_color': '#00008b','text_wrap': True})

worksheet.set_column('A:A',10,red)
worksheet.set_column('B:C',30,format)
worksheet.set_column('D:D',40,format)
worksheet.set_column('E:E',60,format)
worksheet.set_column('F:I',10,format)
worksheet.set_column('J:J',20,yellow)
worksheet.set_column('K:K',15,format)
worksheet.set_column('L:N',10,format)

writer.save()
workbook.close()

del cnxn  # close the connection

# You can also set up Windows desktop notifications for the same.
notifyme("IRIS Usage Tracker - PROD",f"Hi Prateek! \n Total tasks as of {timestr} :\n{n_rows}")

# write an email message
txt = ("Check this out!")

# we will built the message using the email library and send using smtplib
msg = MIMEMultipart()
msg['Subject'] = "Automated report output from Database XYZ from server ABC"  # set email subject
msg.attach(MIMEText(txt))  # add text contents
print(df)
# we will send via outlook, first we initialise connection to mail server
smtp = smtplib.SMTP('smtp-mail.outlook.com', '587')
smtp.ehlo()  # say hello to the server
smtp.starttls()  # we will communicate using TLS encryption

# login to outlook server, using generic email and password
smtp.login('XYZ@outlook.com', 'Password123@')

# send email to our boss
smtp.sendmail('ABC@outlook.com', 'XYZ@outlook.com', msg.as_string())

# finally, disconnect from the mail server
smtp.quit()
