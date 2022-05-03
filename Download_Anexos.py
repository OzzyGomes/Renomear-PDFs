#!/usr/bin/env python
# coding: utf-8

# In[19]:


# import libraries
import win32com.client
import re
import datetime
# set up connection to outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items
message = messages.GetFirst()
today_date = str(datetime.date.today())


# In[ ]:


while True:
    try:
        #current_sender = str(message.Sender).lower()
        #current_subject = str(message.Subject).lower()
        # find the email from a specific sender with a specific subject
        # condition
        if re.search('Subject Title',current_subject) != None and    re.search(sender_name,current_sender) != None:
            print(current_subject) # verify the subject
            print(current_sender)  # verify the sender
            attachments = message.Attachments
            attachment = attachments.Item(1)
            attachment_name = str(attachment).lower()
            attachment.SaveASFile(path + '\\' + attachment_name)
        else:
            pass
        message = messages.GetNext()
    except:
        message = messages.GetNext()
exit


# In[10]:


import datetime
import os
import win32com.client


path = os.path.expanduser("C:/Users/ojgomes/Documents/Anexos_Boletos")  #location o file 
today = datetime.date.today()  # current date if you want current

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")  #opens outlook
inbox = outlook.GetDefaultFolder(6) 
messages = inbox.Items


def saveattachemnts(subject):
    for message in messages:
        if message.Subject == subject and message.Unread:
        #if message.Unread:  #I usually use this because the subject line contains time and it varies over time. So I just use unread

            # body_content = message.body
            attachments = message.Attachments
            attachment = attachments.Item(1)
            for attachment in message.Attachments:
                attachment.SaveAsFile(os.path.join(path, str(attachment)))
                if message.Subject == subject and message.Unread:
                    message.Unread = False
                break                
                
saveattachemnts('teste')
saveattachemnts('EXAMPLE 2')


# In[13]:


message = messages.GetLast()
txt_arq = message.subject
saveattachemnts(txt_arq)


# In[37]:


import win32com.client
import re
import datetime
import os
import email


# In[30]:


directory = "./boletos"

files_in_directory = os.listdir(directory)

for file in files_in_directory:
    if '.pdf' not in file:
        path_to_file = os.path.join(directory, file)
        os.remove(path_to_file)
        


# In[34]:




path = os.path.expanduser("C:/Users/ojgomes/Documents/Automação Rent/Boletos")

Today = datetime.date.today()
today = Today.strftime('%Y%m%d')
 
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
message = inbox.items
 
items = inbox.items
for item in items:
    item.Unread = False
    RT = item.ReceivedTime
    Msgdate = datetime.datetime(RT.year ,RT.month, RT.day, RT.hour, RT.minute, RT.second)
    msgdate = Msgdate.strftime('%Y%m%d')
    if today == msgdate:
        for attachment in item.Attachments:
            #print(attachment.FileName)
            attachment.SaveAsFile(os.path.join(path, str(attachment)))
            

directory = "./boletos"

files_in_directory = os.listdir(directory)

for file in files_in_directory:
    if '.pdf' not in file.lower():
        path_to_file = os.path.join(directory, file)
        os.remove(path_to_file)

