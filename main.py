import win32com.client
import pandas as ps

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
# Assign an Outlook

mailbox = outlook.Folders("Daily Report")
# Assign mailbox

folder = mailbox.Folders("Performance")
# Assign Folder

emails = folder.Items
# Iterate folder

content = []
for email in emails:
    content.append([email.Subject, email.SenderEmailAddress, email.ReceivedTime, email.Body])
    # Insert mail content

df = ps.DataFrame(content, columns=['Subject', 'Sender', 'Date', 'Body'])
df.to_excel("output.xlsx", index=False)