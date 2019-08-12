import win32com.client
import os
import pandas as pd

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case the inbox. You can change that number to reference
messages = inbox.Items
message = messages.GetFirst()
subject = message.Subject

# path of directory where attached files are to be saved locally
get_path = 'I:\\github_repos\\work-8\\download_mail'
# print(len(messages))      # length of the messages in "Inbox" folder

# Loop in all the Inbox messages
for m in messages:
    if m.Subject == "test":     # check if the subject == "test"

        # print (message)
        attachments = message.Attachments
        num_attach = len([x for x in attachments])
        for x in range(1, num_attach+1):
            attachment = attachments.Item(x)
            attachment.SaveASFile(os.path.join(get_path,attachment.FileName))
        # print (attachment)
        message = messages.GetNext()

    else:
        message = messages.GetNext()

df_attachments = []
for dirname, dirpath, files in os.walk("./download_mail/"):
    for file in files:
        # print(file)
        df_attachments.append(pd.ExcelFile("./download_mail/" + file).parse(0))      # parse sheet_name - 0 i.e. 1st sheet

# define `df1` i.e. from downloaded mail attachments (in excel format)
df1 = df_attachments[0]
# print(df1.head())

# define `df2` i.e. output
df2 = pd.ExcelFile("./output/Attachment_1564702282.xlsx").parse(0)              # parse sheet_name - 0 i.e. 1st sheet
# print(df2)


