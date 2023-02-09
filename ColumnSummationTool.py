import tkinter as tk
from tkinter import filedialog as fd
import openpyxl
# from ColSumWindow import ColSumWindow

class ColumnSummationTool:

    def __init__(self):
        # The first workbook
        self.wb1 = ""
        # The second workbook (may be the same as the first)
        self.wb2 = ""
        # Initialize the first and second worksheets. These will be selected from a dropdown of
        # Available worksheets later
        self.ws1 = ""
        self.ws2 = ""

        # # Set up the GUI
        # self.gui = ColSumWindow()




    def openWorkbooks(self, workbookNumber, fileName):
        if workbookNumber == 1:
            self.wb1 = openpyxl.load_workbook(fileName)
        elif workbookNumber==2:
            self.wb2 = openpyxl.load_workbook(fileName)

        print("Loaded workbook {num}".format(num=workbookNumber))


    def addCol(self):
        pass