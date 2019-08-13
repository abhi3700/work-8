import win32com.client
import os
import pandas as pd
from input import *

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case the inbox. You can change that number to reference
messages = inbox.Items
message = messages.GetFirst()
subject = message.Subject


# print(len(messages))      # length of the messages in "Inbox" folder

# Loop in all the Inbox messages
for m in messages:
    if m.Subject == mail_sub:     # check if the subject == "test"

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
for dirname, dirpath, files in os.walk(download_mail_dir):
    for file in files:
        # print(file)
        df_attachments.append(pd.ExcelFile(download_mail_dir + file).parse(0))      # parse sheet_name - 0 i.e. 1st sheet

# define `df1` i.e. from downloaded mail attachments (in excel format)
df1 = df_attachments[0]
# print(df1.head())

# define `df2` i.e. output
df2 = pd.ExcelFile(output_file).parse(0)              # parse sheet_name - 0 i.e. 1st sheet
# print(df2)

"""
Map 2 column names: df1 --> df2
===============================
df1                     df2
-----------------------------------------
"Order ID"              "ClientOrderCode"  
"Billing First Name"    "FirstName"
"Shipping Address 1"    "Address1"
"Shipping City"         "City"
"Shipping State"        "Province"
"Shipping Postcode"     "PostCode"
"Shipping Country"      "Country"
"SKU"                   "OrderSourceSKU"
"Purchased"             "ProductNum"
"""
col_1 = ["Order ID", "Billing First Name", "Shipping Address 1", "Shipping City",  "Shipping State", 
        "Shipping Postcode", "Shipping Country", "SKU", "Purchased"]
col_2 = ["ClientOrderCode", "FirstName", "Address1", "City", "Province", "PostCode", "Country", "OrderSourceSKU", "ProductNum"]

# define `df1_col1` for `df1` as per the required columns i.e. col_1
df1_col1 = df1[col_1]
# print(df1_col1.head())

# define `df2_col2` for `df2` as per the required columns i.e. col_2
df2_col2 = df1_col1             # copy the content from `df1_col1` into `df2_col2`
df2_col2.columns = col_2        # add columns
# print(df2_col2.head())

# drop all rows of `df2`
df2.drop(df2.index, inplace=True)

"""
NOTE: Before applying this, `df2` needs to be cleared first.
and then copy respective columns data into from `df2_col2` into `df2`
"""
for c in col_2:
    df2[c] = df2_col2[c]

# print(df2)

# print the `df2` (w/o index) into already present excel file 
df2.to_excel(output_file, index= False)

