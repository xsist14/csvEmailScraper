# to do 
# 1 grab email content
# 2 specify what to grab
# put content in csv file

import smtplib
import time
import imaplib
import email
import re
import csv

# -------------------------------------------------
#
# Utility to read email from Gmail Using Python
#
# ------------------------------------------------


ORG_EMAIL   = "@rgrmarketing.com"
FROM_EMAIL  = "<username>" + ORG_EMAIL
FROM_PWD    = "<password>"
SMTP_SERVER = "imap-mail.outlook.com"
SMTP_PORT   = 993




def read_email_from_outlook():
    with open('mycsv.csv', 'w') as f:
        thewriter = csv.writer(f)
        thewriter.writerow(['qerror', 'sql', 'end'])
        try:
            mail = imaplib.IMAP4_SSL(SMTP_SERVER)
            mail.login(FROM_EMAIL,FROM_PWD)
            mail.select('database errors')

            type, data = mail.search(None, 'ALL')
            mail_ids = data[0]

            id_list = mail_ids.split()   
            first_email_id = int(id_list[0])
            latest_email_id = int(id_list[-1])


            for i in range(2418,latest_email_id, 1):
                typ, data = mail.fetch(i, '(RFC822)' )

                for response_part in data:
                    if isinstance(response_part, tuple):
                        msg = email.message_from_string(response_part[1])
                        email_to = msg['to']
                        email_from = msg['from']
                        email_subject = msg['subject']
                        # print 'To : ' + email_to + '\n'
                        # print 'From : ' + email_from + '\n'
                        # print 'Subject : ' + email_subject + '\n'
                        if msg.is_multipart():
                            for payload in msg.get_payload():
                                # if payload.is_multipart(): ...
                                print payload.get_payload()
                        else:
                            csvContent = msg.get_payload()
                            arr = csvContent.split("\r\n\r----------------------\r\n\r")
                            # print(arr[7])
                            # print(arr[8])
                            # print(arr[9])
                            errorMessage = arr[7]
                            sqlQuery = arr[8]
                            endPoint = arr[9]
                            thewriter.writerow([errorMessage, sqlQuery, endPoint])

                            # print('\n' + '\n' + "This is an error message: " + errorMessage + '\n' + '\n')
                            # print("This is a sql query: " + sqlQuery + '\n' + '\n')
                            # print("This is an endpoint: " + endPoint + '\n' + '\n')
                            # print("last email id: ")
                            # print(latest_email_id) 
                            # print("first email id: ")
                            # print(first_email_id)

                        # sleep(5)

        except Exception, e:
            print str(e)

read_email_from_outlook()