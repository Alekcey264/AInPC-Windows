from time import sleep
from PyQt6.QtCore import QThread, pyqtSignal
from PyQt6.QtWidgets import QLabel
from dll_fetch import *

class WorkerThread(QThread):
    mysignal = pyqtSignal(str)
    def __init__(self, parent = None):
        QThread.__init__(self, parent)

    def run(self, handle, root, text, list):
        '''result = fetch_stats(handle, root, text, list)
        self.mysignal.emit(result)'''
        pass

class HyperlinkLabel(QLabel):
    def __init__(self, text, link):
        super().__init__()
        self.setText(f'<a href="{link}">{text}</a>')
        self.setOpenExternalLinks(True)