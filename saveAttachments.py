import datetime
import os
import win32com.client


path = os.path.expanduser("~/Desktop/Attachments")
today = datetime.date.today()

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6).Folders("Flir images")
messages = inbox.Items


def saveattachemnts(subject):
    for message in reversed(messages):
        if message.Subject == subject and message.Unread or message.ReceivedTime.date() == today:
            # body_content = message.body
            attachments = message.Attachments
            attachment = attachments.Item(1)
            for attachment in message.Attachments:
                f, ext = os.path.splitext(str(attachment))
                f = f + '_' + message.ReceivedTime.strftime("%Y-%m-%d_%H%M%S") + ext
                print(f)
                attachment.SaveAsFile(os.path.join(path, f))
                if message.Subject == subject and message.Unread:
                    message.Unread = False


saveattachemnts('<no subject>')
