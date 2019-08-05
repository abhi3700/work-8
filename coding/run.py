import win32com.client
import os
import xlwings

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case the inbox. You can change that number to reference
messages = inbox.Items
message = messages.GetFirst()
subject = message.Subject

# path of directory where attached files are to be saved locally
get_path = 'C:\\Users\\abhijit\\Desktop\\mail'
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