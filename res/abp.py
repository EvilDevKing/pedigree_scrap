import sys, os, base64
from multiprocessing import Process
from constants import getGoogleService, createFileWith, ORDER_DIR_NAME

import thread1 as t1
import thread2 as t2
import thread3 as t3
import thread4 as t4

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

def run(sheetId, sheetName):
    sys.stdout = Unbuffered(sys.stdout)

    print("Preparing to run script...")
    pdf_cnt = 0
    service = getGoogleService('gmail', 'v1')
    results = service.users().messages().list(userId='me', labelIds=['INBOX'], q="from:noreply@aqha.org").execute()
    messages = results.get('messages')
    if messages:
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
                        pdf_cnt += 1
                        createFileWith(ORDER_DIR_NAME + "/" + filename, file_data, 'wb')
                service.users().messages().delete(userId='me', id=msg_id).execute()
            except: continue

    if os.path.exists("res/t1.txt"):
        os.remove("res/t1.txt")

    print("Preparing done.")
    proc1 = Process(target=t1.start, args=[sheetId, sheetName, pdf_cnt])
    proc2 = Process(target=t2.start, args=[pdf_cnt])
    proc3 = Process(target=t3.start, args=[sheetId, sheetName, pdf_cnt])
    
    proc1.start()
    proc2.start()
    proc3.start()
    
    proc1.join()
    proc2.join()
    proc3.join()

    print("""
        ######################################
        ######################################      
        ######################################
        ######################################
    """)

    proc2 = Process(target=t2.start)
    proc4 = Process(target=t4.start, args=[sheetId, sheetName])

    proc2.start()
    proc4.start()

    proc2.join()
    proc4.join()
    
    print("Done!")
    
if __name__ == "__main__":
    sheetId = input("Enter your worksheet id: ")
    if sheetId.strip() == "":
        print("The sheetid can't be empty.")
    else:
        sheetName = input("Enter your worksheet name: ")
        run(sheetId, sheetName)