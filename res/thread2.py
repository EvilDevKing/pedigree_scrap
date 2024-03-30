import base64, os, time, sys
from constants import *

class Unbuffered(object):
   def __init__(self, stream):
       self.stream = stream
   def write(self, data):
       self.stream.write(data)
       self.stream.flush()
   def writelines(self, datas):
       self.stream.writelines(datas)
       self.stream.flush()
   def __getattr__(self, attr):
       return getattr(self.stream, attr)

def checkMailAndDownloadOrderFile():
    # Create Gmail API service
    service = getGoogleService('gmail', 'v1')
    pdf_cnt = 0
    while True:
        if os.path.exists("res/t3.txt"):
            os.remove("res/t3.txt")
            for file in getOrderFiles():
                os.remove(ORDER_DIR_NAME + "/" + file)
            break
        if os.path.exists("res/t4.txt"):
            os.remove("res/t4.txt")
            for file in getOrderFiles():
                os.remove(ORDER_DIR_NAME + "/" + file)
            break
        # Fetch messages from inbox
        results = service.users().messages().list(userId='me', labelIds=['INBOX'], maxResults=10, q="from:noreply@aqha.org").execute()
        messages = results.get('messages')
        if not messages:
            print("No messages found.")
        else:
            print("New Messages: " + str(len(messages)))
            for message in messages:
                msg_id = message['id']
                msg_body = service.users().messages().get(userId='me', id=msg_id).execute()
                try:
                    parts = msg_body['payload']['parts']
                    for part in parts:
                        if part['filename']:
                            if 'data' in part['body']:
                                data = part['body']['data']
                            else:
                                att_id = part['body']['attachmentId']
                                att = service.users().messages().attachments().get(userId='me', messageId=msg_id, id=att_id).execute()
                                data = att['data']
                            file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
                            filename = part['filename']
                            
                            createFileWith(ORDER_DIR_NAME + "/" + filename, file_data, 'wb')
                            pdf_cnt += 1
                            print("Stored a pdf order file : " + filename)
                    service.users().messages().delete(userId='me', id=msg_id).execute()
                except: continue
        time.sleep(5)

def start():
    sys.stdout = Unbuffered(sys.stdout)
    print("Second process started")
    createOrderDirIfDoesNotExists()
    checkMailAndDownloadOrderFile()
    print("Second process finished")