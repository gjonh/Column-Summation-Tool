import tkinter as tk
from tkinter import filedialog as fd
import openpyxl
# from ColSumWindow import ColSumWindow

class ColumnSummationTool:

    def __init__(self):
        # The first workbook
        self.wb1 = ""
        self.wb1Name=""
        # The second workbook (may be the same as the first)
        self.wb2 = ""
        self.wb2Name=""
        # Initialize the first and second worksheets. These will be selected from a dropdown of
        # Available worksheets later
        self.ws1 = ""
        self.ws2 = ""

        # # Set up the GUI
        # self.gui = ColSumWindow()




    def openWorkbooks(self, workbookNumber, fileName):
        if workbookNumber == 1:
            self.wb1 = openpyxl.load_workbook(fileName)
            self.wb1Name = fileName
        elif workbookNumber==2:
            self.wb2 = openpyxl.load_workbook(fileName)
            self.wb2Name = fileName

        print("Loaded workbook {num}".format(num=workbookNumber))


    def setSheet(self, sheet1, sheet2):
        sheet1=str(sheet1)
        sheet2=str(sheet2)
        returnCode=True
        try:
            self.ws1 = self.wb1[sheet1]
            print("Loaded {sheet1} from workbook 1".format(sheet1=sheet1))
        except:
            print("Could not open sheet {sheet1} in workbook 1".format(sheet1=sheet1))
            returnCode=False
        try:
            self.ws2 = self.wb1[sheet2]
            print("Loaded {sheet2} from workbook 2".format(sheet2=sheet2))
        except:
            print("Could not open sheet {sheet2} in workbook 2".format(sheet2=sheet2))
            returnCode=False
        return returnCode

    def addColumnValues(self, col1, col2):
        sum1 = 0.0
        sum2 = 0.0
        returnCode = True
        position = 0
        for cell in self.ws1[col1]:
            position+=1
            try:
                value = cell.value
                value = float(value)
                sum1+=value
            except:
                print("[First Column] Non-number value at row {row} with value {value}".format(row=position, value=value))
        for cell in self.ws2[col2]:
            position+=1
            try:
                value = float(cell.value)
                sum2+=value
            except:
                print("[First Column] Non-number value at row {row} with value {value}".format(row=position, value=value))



        return sum1, sum2



