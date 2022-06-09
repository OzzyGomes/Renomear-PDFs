import win32com.client
import datetime
import os


path = os.path.expanduser(os.path.join(os.getcwd(), "Boletos").replace('\\',"/"))

if not os.path.exists(path):
    os.makedirs(path)


Today = datetime.date.today()
today = Today.strftime('%Y%m%d')
 
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
message = inbox.items
 
items = inbox.items

folder_PDFs = inbox.Folders['Emails_PDFs']


def createReply(email):
        reply = email.Reply()
        newBody = """
        <style>
        p {
            font-size: 16px;
        }
        </style>
        
        <p>Olá,</p>
        
        <p>Arquivo recebido com sucesso.</p>
        
            
        """
        reply.HTMLBody = newBody + reply.HTMLBody
        reply.Send()
        


for item in items:
    
    RT = item.ReceivedTime
    Msgdate = datetime.datetime(RT.year ,RT.month, RT.day, RT.hour, RT.minute, RT.second)
    msgdate = Msgdate.strftime('%Y%m%d')
    
    if today == msgdate:
        #item.Unread = True
        for attachment in item.Attachments:
            if attachment.FileName.lower().endswith('.pdf'):
                attachment.SaveAsFile(os.path.join(path, str(attachment)))
                createReply(item)
                item.Move(folder_PDFs)
                item.Unread 
                print(attachment.FileName)
                
                



