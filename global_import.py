import sys, time, os, subprocess
from PyQt6.QtCore import Qt, QSize, QTimer, QThread, pyqtSignal, QObject
from PyQt6.QtWidgets import QTableView, QScrollArea, QAbstractItemView, QHeaderView, QTreeWidgetItem, QTreeWidget, QHBoxLayout, QListWidget, QTableWidgetItem, QTableWidget, QLabel, QLineEdit, QApplication, QMainWindow, QTabWidget, QWidget, QGridLayout, QVBoxLayout, QFormLayout, QMessageBox
from PyQt6.QtGui import QPalette, QIcon, QAction, QColor
import pyqtgraph, numpy, sqlite3
from pyuac import main_requires_admin
from pywintypes import error as PyWinError