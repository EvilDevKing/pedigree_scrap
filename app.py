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

import traceback

try:
    from ctypes import windll  # Only exists on Windows.
    myappid = 'BrittanyHoly.Inbreeding.Pedigree'
    windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except ImportError:
    pass

class WorkerSignals(QObject):
    finished = pyqtSignal()
    error = pyqtSignal(tuple)
    result = pyqtSignal(object)
    progress = pyqtSignal(int)

class ThreadWorker1(QRunnable):
    def __init__(self, fn, *args, **kwargs):
        super(ThreadWorker1, self).__init__()

        # Store constructor arguments (re-used for processing)
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.signals = WorkerSignals()

        # Add the callback to our kwargs
        self.kwargs['progress_callback'] = self.signals.progress

    @pyqtSlot()
    def run(self):
        # Retrieve args/kwargs here; and fire processing using them
        try:
            result = self.fn(*self.args, **self.kwargs)
        except:
            traceback.print_exc()
            exctype, value = sys.exc_info()[:2]
            self.signals.error.emit((exctype, value, traceback.format_exc()))
        else:
            self.signals.result.emit(result)  # Return the result of the processing
        finally:
            self.signals.finished.emit()  # Done

class MainWindow(QMainWindow):
    def __init__(self, *args, **kwargs):
        super(MainWindow, self).__init__(*args, **kwargs)

        self.basedir = os.path.dirname(__file__)
        self.pdf_thread = None
        self.abp_proc_one = None
        self.abp_proc_two = None
        self.abp_proc_three = None

        self.setWindowTitle('Pedigree Application')
        self.setWindowIcon(QIcon(os.path.join(self.basedir, "icon.ico")))
        self.setFixedSize(QSize(800, 600))

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

        self.setCentralWidget(self.tab)
        self.show()
        
        self.threadpool = QThreadPool()
        
    def log(self, viewId, msg):
        if viewId == "pdf":
            self.logview_one.append("<span style='color: \"white\"'>" + msg + "</span>")
        else:
            self.logview_two.append("<span style='color: \"white\"'>" + msg + "</span>")
    
    def onPdfButtonClicked(self):
        self.logview_one.clear()
        self.pdf_thread = Thread(target=pdfMain.run, args=[self,])
        self.pdf_thread.start()
        
    def onAllButtonClicked(self):
        self.start_button_two.setText("Processing...")
        self.start_button_two.setEnabled(False)
        self.start_button_one.setEnabled(False)
        self.logview_two.clear()
        
        self.abp_proc_one = Process(target=t1.start, args=[self.child_conn])
        # self.abp_proc_two = Process(target=t2.start, args=[self.output_queue])
        # self.abp_proc_three = Process(target=t3.start, args=[self.output_queue])
        
        self.abp_proc_one.start()
        # self.abp_proc_two.start()
        # self.abp_proc_three.start()
        
        self._timer = QTimer()
        self._timer.setInterval(1000)
        self._timer.timeout.connect(self.display_results)
        self._timer.start()
        
        self.abp_proc_one.join()
        # self.abp_proc_two.join()
        # self.abp_proc_three.join()
        
        self.start_button_two.setText("Start")
        self.start_button_two.setEnabled(True)
        self.start_button_one.setEnabled(True)
    
    def display_results(self):
        print(self.parent_conn.recv())
        self.logview_two.append("<span style='color: white;'>" + self.parent_conn.recv() + "</span>")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    
    ret = app.exec_()
    sys.exit(ret)