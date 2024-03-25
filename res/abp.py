import sys
from multiprocessing import Process

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
    
    proc1 = Process(target=t1.start, args=[sheetId, sheetName])
    proc2 = Process(target=t2.start)
    proc3 = Process(target=t3.start, args=[sheetId, sheetName])
    
    proc1.start()
    proc2.start()
    proc3.start()
    
    proc1.join()
    proc2.join()
    proc3.join()

    proc4 = Process(target=t4.start, args=[sheetId, sheetName])
    proc4.start()
    proc4.join()
    
    print("Done!")
    
if __name__ == "__main__":
    sheetId = input("Enter your worksheet id: ")
    if sheetId.strip() == "":
        print("The sheetid can't be empty.")
    else:
        sheetName = input("Enter your worksheet name: ")
        run(sheetId, sheetName)