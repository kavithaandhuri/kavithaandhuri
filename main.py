import os
from datetime import datetime, timedelta
import win32com.client
import pandas as pd
def send_mail(df, message):
    outlook = win32com.client.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.Subject = 'Test code for Follow up'
    mail.To = "kavitha.andhuri@capgemini.com"
    # html_table_1, html_table_2 = fetch_sql()
    html_style = '<style>table{border: 1px solid black;border-collapse: collapse; width=50%;}th{ ' \
                 'background-color:#437d88;}</style> '

    mail.HTMLBody = '<html><head>' + html_style + '</head><body>Hai, <br><br>' + message + '<br>' + df + \
                    '<br><br> Thanks <br>Kavitha</body></html> '
    '''  # time_format = datetime.pastime("2022-07-25 22:15:00 +00:00", '%Y-%m-%d %H:%M:%S %z')'''
    mail.Send()

if __name__ == '__main__':
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)
    messages = inbox.Items
    subject = "Synq Downtime Notification Schedule"  # replace with desired subject text
    # filter emails based on subject
    emails = [message for message in messages if message.Subject == subject]
    # create a list of dictionaries to store email information
    data = []
    for email in emails:
        data.append({"Subject": email.Subject,
                     "Sender": email.Sender.Name,
                     "Body": email.Body,
                     "status": "received downtime schedule"})

    # create a pandas DataFrame from the email data
    df = pd.DataFrame(data)

    # save the DataFrame as an Excel file
    send_mail(df.to_html(index=False), "PFA")
    for j in df.index:
        df.loc[j, 'status'] = 'Sent to IKEA System'
    df.to_excel("emails1.xlsx", index=False)
