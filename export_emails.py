import win32com.client as client
import csv
import json
from datetime import datetime


def mail_search(term, sub_, folder):
    """Recursively search all folders for mails containing search term"""
    relevant_messages = [(item, item.Parent.Name) for item in folder.Items if term in item.Body.lower()]

    if sub_ == "True":
        # check for subfolders (base case)
        subfolder_count = folder.Folders.Count
        # search all subfolders
        if subfolder_count > 0:
            for subfolder in folder.Folders:
                relevant_messages.extend(mail_search(term, sub_, subfolder))

    return relevant_messages


# Press the green button in the gutter to run the script.
if __name__ == '__main__':

    # today = datetime.now()

    with open("export_configuration.json", "r") as mailbox_file:
        creds = json.load(mailbox_file)x`
        sub_ = creds['mailbox']['subs']
        term = creds['mailbox']['term']
        duration = creds['mailbox']['from_date']
        duration = datetime.strptime(duration, '%Y-%m-%d %H:%M:%S.%f%z')

    outlook = client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    recip = outlook.CreateRecipient(creds['mailbox']['account'])
    inbox = outlook.GetSharedDefaultFolder(recip, 6)
    results = mail_search(term, sub_, inbox)

    with open(creds['mailbox']['filepath'], 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['ParentFolder', 'Subject', 'ReceivedTime', 'From', 'Categories'])
        for message, parent in results:
            if datetime.strptime(str(message.ReceivedTime), '%Y-%m-%d %H:%M:%S.%f%z') >= duration:
                writer.writerow([parent,
                                 message.Subject,
                                 message.ReceivedTime,
                                 message.SenderName,
                                 message.Categories])
