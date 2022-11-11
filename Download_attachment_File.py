'''
Author       :   Anuraj Pilanku(CAC Automation Developer- ADM MLEU Legacy)

Code Utility : Inspecting Outlook is opened or not , if not open it . Connect to outlook and navigate to required Folder. Collect Mail Details and Then Download Attachment.
                Finally Close\Kill\Terminate Outllook Application

Version      :  1.0

Created Date : 11-11-2022

'''

import win32com.client
from win32com.client import Dispatch
import win32ui
import psutil,os
from pywinauto import Desktop

'''Inspect Outlook is opened or not , if not open outlook'''

def outlook_is_running():
    import win32ui
    try:
        win32ui.FindWindow(None, "Microsoft Outlook")
        return {"status":True,"update":"outlook is already opened"}
    except win32ui.error:
        return {"status":False,"update":"Error with win32ui module"}
def postprocedure():
    output=dict()
    execution=outlook_is_running()
    status=execution["status"]
    output["update"]=execution["update"]
    #print(execution["update"])
    if not status:#outlook_is_running():
        import os
        os.startfile("outlook")
        output["process"] = "Opening Outlook..!!"
        #print("Opening Outlook..!!")
    return output
run=postprocedure()
print(run)

'''
Connecting to Mail Folders
'''

outlook=Dispatch("Outlook.Application").GetNamespace("MAPI")

foldername='Sent Items'
subjectkeyword="Applens"
senders=["P.Anuraj@cognizant.com"]
attachmentDownloadPath="C:\\Users\\2040664\\anuraj\\lcs_borker"
def downloadAttachment(foldername,subjectkeyword,senders,attachmentDownloadPath):
    emaillist=[]
    try:
        Smail_id=dict()
        root_folder=outlook.Folders.Item(1)
        #print(root_folder.Name)
        '''Get All Folder names'''
        #for folder in root_folder.Folders:
            #print(folder.Name)
        '''accessing_mail_folder'''
        requiredFolder=root_folder.Folders[foldername]
        '''For accesing Folder inside Folder'''
        # subFolder=outlook.Folders["PrimaryFoldername"].Folders["InnerFoldername"].Folders["ChildFoldername"]
        #print(requiredFolder)
        for message in requiredFolder.Items:
            #if message.SenderEmailType == "EX":--print("Email Address: ", message.Sender.GetExchangeUser().PrimarySmtpAddress)
            #elif message.SenderEmailType == "SMTP":--- print("Email Address: ", message.SenderEmailAddress)
            #print(message)
            if message.UnRead==True:
                if message.SenderEmailType == "EX":
                    SendermailID=str(message.Sender.GetExchangeUser().PrimarySmtpAddress).lower().strip()
                elif message.SenderEmailType == "SMTP":
                    SendermailID=message.SenderEmailAddress
                if subjectkeyword.strip().lower() in str(message.Subject).strip().lower() :
                    #if str(message.Sender.GetExchangeUser().PrimarySmtpAddress).lower().strip() in list(map(str.lower,senders)):
                    if SendermailID in list(map(str.lower, senders)):
                        message.UnRead=False#change mail from unread state to read state
                        #print(message.Subject)
                        path = attachmentDownloadPath
                        subject = message.Subject
                        body=str(message.Body).replace('\r',"").replace("\n","")
                        recipients=message.Recipients
                        for recipient in recipients:
                            address=str(recipient.AddressEntry.Address.lower()).replace('\r',"").replace("\n","")
                            emaillist.append(address)
                        #print(subject)
                        sender = str(message.Sender)
                        #sendermailid = str(message.SenderEmailAddress)
                        try:
                            attachments = message.Attachments  # returns All attachments in the mail
                            #requiredAttachment = attachments.Item(1)
                            #attachmentName = str(requiredAttachment).lower()
                            #requiredAttachment.SaveASFile(path + '\\' + attachmentName)
                            num_attach=len([x for x in attachments])
                            if num_attach >=1:
                                for x in range(1,num_attach+2):
                                    attachment=attachments.Item(x)
                                    if (attachment.FileName).endswith('xlsx'):
                                        attachment.SaveASFile(path + '\\' + attachment.FileName)
                                        return {"rootFolder":root_folder.Name,"mailsubject":subject,"mailsender":sender,"FileType":"Excel","attachmentName":attachment.FileName,"Recipients":emaillist,"Body":body}
                                    elif (attachment.FileName).endswith('pdf'):
                                        attachment.SaveASFile(path + '\\' + attachment.FileName)
                                        return {"rootFolder":root_folder.Name,"mailsubject":subject,"mailsender":sender,"FileType":"PDF","attachmentName":attachment.FileName,"Recipients":emaillist,"Body":body,"mailid":str(message.Sender.GetExchangeUser().PrimarySmtpAddress).lower().strip()}
                                    elif (attachment.FileName).endswith('txt'):
                                        attachment.SaveASFile(path + '\\' + attachment.FileName)
                                        return {"rootFolder":root_folder.Name,"mailsubject":subject,"mailsender":sender,"FileType":"NotePad","attachmentName":attachment.FileName,"Recipients":emaillist,"Body":body}
                                    else:
                                        attachment.SaveASFile(path + '\\' + attachment.FileName)
                                        return {"rootFolder": root_folder.Name, "mailsubject": subject, "mailsender": sender,"Recipients":emaillist,
                                                "FileType": "NotePad","attachmentName":attachment.FileName,"Body":body}
                            else:
                                return {"rootFolder": root_folder.Name, "mailsubject": subject, "mailsender": sender,"Attachment":"Attachment_Absent","Recipients":emaillist,"Body":body}
                        except Exception as e:
                            return "Mail Attachment Error: "+str(e)

    except Exception as e:
        return "Error: " + str(e)
download=downloadAttachment(foldername,subjectkeyword,senders,attachmentDownloadPath)
print(download)

''' Closing Outlook Application'''

def closeOutlook():
    windows = Desktop(backend="uia").windows()
    windowTitle=[w.window_text() for w in windows]
    for title in windowTitle:
        if "Outlook" in title:
            terminate=os.system("TASKKILL /F /IM OUTLOOK.EXE")
            return "Outlook Application has been successfully terminated., "+str(terminate)
        else:
            return "Outlook Application Already Closed"
closeOutlook()
