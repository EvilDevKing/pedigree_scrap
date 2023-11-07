import sys, os
from threading import *
from multiprocessing import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *

from pdfscript import main as pdfMain
from allbreedpedigree import thread1 as t1
from allbreedpedigree import thread2 as t2
from allbreedpedigree import thread3 as t3

try:
    from ctypes import windll  # Only exists on Windows.
    myappid = 'BrittanyHoly.Inbreeding.Pedigree'
    windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except ImportError:
    pass

class MainWindow(QMainWindow):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)

        self.basedir = os.path.dirname(__file__)
        self.pdf_thread = None
        self.abp_recv_one = None
        self.abp_send_one = None
        self.abp_recv_two = None
        self.abp_send_two = None
        self.abp_recv_three = None
        self.abp_send_three = None
        
        self.abp_proc_one = None
        self.abp_proc_two = None
        self.abp_proc_three = None

        self.setWindowTitle('Pedigree Application')
        self.setWindowIcon(QIcon(os.path.join(self.basedir, "icon.ico")))
        self.SetFixedSize(QSize(800, 600))

        self.layout = QVBoxLayout()
        self.layout.setStretch(1, 1)

        # create a tab widget
        self.tab = QTabWidget()

        # pdf script page
        self.pdf_page = QWidget()
        self.layout_one = QVBoxLayout()
        self.start_button_one = QPushButton("Start")
        self.start_button_one.setFixedSize(QSize(150, 50))
        self.start_button_one.clicked.connect(self.onPdfButtonClicked)
        
        self.label_output = QLabel("Output:", self)
        
        self.logview_one = QTextEdit()
        self.logview_one.setReadOnly(True)
        self.logview_one.setLineWrapMode(QTextEdit.LineWrapMode.WidgetWidth)
        self.logview_one.setStyleSheet(
        """
        QTextEdit{
            background: rgb(50, 50, 50);
        }
        """
        )
        
        self.layout_one.addWidget(self.start_button_one, alignment=Qt.AlignmentFlag.AlignCenter)
        self.layout_one.addWidget(self.label_output)
        self.layout_one.addWidget(self.logview_one)
        self.pdf_page.setLayout(self.layout_one)

        # allbreedpedigree script page
        self.abp_page = QWidget()
        self.layout_two = QVBoxLayout()
        self.start_button_two = QPushButton("Start")
        self.start_button_two.setFixedSize(QSize(150, 50))
        self.start_button_two.clicked.connect(self.onAllButtonClicked)
        
        self.label_output = QLabel("Output:", self)
        
        self.logview_two = QTextEdit()
        self.logview_two.setReadOnly(True)
        self.logview_two.setLineWrapMode(QTextEdit.LineWrapMode.WidgetWidth)
        self.logview_two.setStyleSheet(
        """
        QTextEdit{
            background: rgb(50, 50, 50);
        }
        """
        )
        
        self.layout_two.addWidget(self.start_button_two, alignment=Qt.AlignmentFlag.AlignCenter)
        self.layout_two.addWidget(self.label_output)
        self.layout_two.addWidget(self.logview_two)
        self.abp_page.setLayout(self.layout_two)

        # add pane to the tab widget
        self.tab.addTab(self.pdf_page, 'PDF Script')
        self.tab.addTab(self.abp_page, 'All Breed Pedigree Script')
        
        self.layout.addWidget(self.tab)

        self.setLayout(self.layout)
        self.setWindowTitle("Pedigree Application")
        self.setWindowIcon(QIcon('icon.ico'))
        self.setFixedSize(800, 600)
        
        self.show()
        
    def log(self, viewId, msg):
        if viewId == "pdf":
            self.logview_one.append("<span style='color: \"white\"'>" + msg + "</span>")
        else:
            self.logview_two.append("<span style='color: \"white\"'>" + msg + "</span>")
    
    def onPdfButtonClicked(self):
        self.logview_one.clear()
        self.pdf_thread = Thread(target=pdfMain.run, args=[self,])
        self.pdf_thread.start()
        
    def startProcOne(parent, pipe):
        t1.start(parent, pipe)
    
    def startProcTwo(parent, pipe):
        t2.start(parent, pipe)
    
    def startProcThree(parent, pipe):
        t3.start(parent, pipe)
        
    def onAllButtonClicked(self):
        self.start_button_two.setText("Processing...")
        self.start_button_two.setEnabled(False)
        self.start_button_one.setEnabled(False)
        self.logview_two.clear()
        
        self.abp_recv_one, self.abp_send_one = Pipe(False)
        self.abp_recv_two, self.abp_send_two = Pipe(False)
        self.abp_recv_three, self.abp_send_three = Pipe(False)
        
        self.abp_proc_one = Process(target=self.startProcOne, args=[self, self.abp_send_one])
        self.abp_proc_two = Process(target=self.startProcTwo, args=[self, self.abp_send_two])
        self.abp_proc_three = Process(target=self.startProcThree, args=[self, self.abp_send_three])
        
        self.abp_proc_one.start()
        self.abp_proc_two.start()
        self.abp_proc_three.start()
        
        self.abp_proc_one.join()
        self.abp_proc_two.join()
        self.abp_proc_three.join()
        
        self.dsp = TextEditDispatcher()
        self.dsp.add_connection(self.logview_two, self.abp_recv_one)
        self.dsp.add_connection(self.logview_two, self.abp_recv_two)
        self.dsp.add_connection(self.logview_two, self.abp_recv_three)
        self.dsp.start()
        
        self.start_button_two.setText("Start")
        self.start_button_two.setEnabled(True)
        self.start_button_one.setEnabled(True)
        
class PipeOutput(object):
    def __init__(self, pipe):
        self.pipe = pipe
    def write(self, s):
        self.pipe.send(s)
    def flush(self):
        pass

class TextEditDispatcher(QThread):
    def __init__(self):
        QThread.__init__(self)
        self.connections = []

    def add_connection(self, widget, pipe):
        self.connections.append((widget, pipe))

    def run(self):
        while (True):
            for widget, pipe in self.connections:
                if pipe.poll():
                    text = pipe.recv().strip()
                    QMetaObject.invokeMethod(widget, 'append', Q_ARG(str, text))
    

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    
    ret = app.exec_()
    if window.abp_proc_one != None:
        window.abp_proc_one.terminate()
    if window.abp_proc_two != None:
        window.abp_proc_two.terminate()
    if window.abp_proc_three != None:
        window.abp_proc_three.terminate()
    sys.exit(ret)