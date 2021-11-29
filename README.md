# Automatic-Python-based-SQL-queries-execution [![Project Status: Concept](https://www.repostatus.org/badges/latest/concept.svg)](https://www.repostatus.org/#concept) [![](https://img.shields.io/badge/Prateek-Ralhan-brightgreen.svg?colorB=ff0000)](https://prateekralhan.github.io/)


As a part of my daily workload, I am supposed to execute lengthy SQL queries on servers,databases as well. Now, this is kind of a really boring work and I thought why not let us try to automate this using Python?? 

And here it is!!

This python script will fetch the results from the Database of your server as per your SQL Query fed into the script and then automatically generate an excel spreadsheet in the folder ***Output_files*** with the pulled out results. I have also formatted the spreadsheet a bit to make it look a little funky :laughing: 
Apart from this, you will also receive a desktop notification for the same as well as get intimated by E-mail. Sounds cool?

## Dependencies:
1. pyodbc
2. pandas
3. os
4. smtplib
5. plyer
6. xlsxwriter

## Installation: 
Our old school ***pip install < package name >*** is always at our service :wink: .

## Usage:
1. Set up your server and Database credentials in the script ***main.py*** in the part ***cnxn_str***.

2. Establish connection with the server and pull results from the database using your SQL Query. You need to append your entire SQL query in the ***query*** string defined in the code. I have used multiline strings here so that you can also include lengthy queries here.

3. Append the results in the Excel spreadsheet. I have formatted it a bit. Please feel free to play around and modify the spreadsheet as per the type of results you are fetching from your server. The spreadsheet that I have in the folder is dummy, I have erased the fetched output due to presence of Sensitive data in that. 

4. Send Desktop notifications for the same - I am using [***plyer***](https://plyer.readthedocs.io/en/latest/) here to generate Desktop notifications. Play around with it and configure as per your choice.

**Note: This feature works only for Win10 OS.** 

5. Set up your E-mail configuration for sending out E-mails after generating the results. I couldn't fully use this technique due to data privacy and openSSL certification issues in my company this isn't allowed. But feel free to test it out and raise an issue if you run into any trouble. You can also check [this](https://www.tutorialspoint.com/send-mail-with-attachment-from-your-gmail-account-using-python) out for sending your Excel sheet as attachment in the Email.

6. Edit your Directory details for this script in the attached batch file : ***sql_python_automate.bat***. We will use this to run the script periodically to receive updates from our database at regular intervals.

7. Use **Windows Task Scheduler** to run this script periodically. Please check out my other work [here](https://github.com/prateekralhan/COVID-19-Tracker-using-Desktop-Notifications-) in case you wish to know how to set up a task in Windows Task Scheduler.
