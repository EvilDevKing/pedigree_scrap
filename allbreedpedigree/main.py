import sys, os
from multiprocessing import Process

import thread1 as t1
import thread2 as t2
import thread3 as t3

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

def run():
    
    sys.stdout = Unbuffered(sys.stdout)
    
    proc1 = Process(target=t1.start)
    proc2 = Process(target=t2.start)
    proc3 = Process(target=t3.start)
    
    proc1.start()
    proc2.start()
    proc3.start()
    
    proc1.join()
    proc2.join()
    proc3.join()
    
    print("Script has worked successfully!")
    
if __name__ == '__main__':
    run()