import tkinter as tk
from tkinter import filedialog as fd
from ColumnSummationTool import ColumnSummationTool


class ColSumWindow:
    def __init__(self):
        # If values are set correctly, allowing you to save to file
        self.canSaveToFile = False

        # This section initializes the tkinter window
        self.root = tk.Tk()

        self.root.title("Summation Verification Tool")
        # WidthxHeight
        self.root.geometry("800x150")
        self.wb1Button = tk.Button(self.root, width=20, text="Select Workbook 1", bg="#ffffff",
                              command=lambda: self.selectWB(workbookNumber=1))
        self.wb2Button = tk.Button(self.root, width=20, text="Select Workbook 2", bg="#ffffff",
                              command=lambda: self.selectWB(workbookNumber=2))
        self.infoLabel = tk.Label(self.root, text="Set both sheets and columns before starting", bg="#FA8A6F")
        self.setButton=tk.Button(self.root, text="Set Values", bg="#32ed35", width=20, command=lambda: self.setVals())
        self.startButton=tk.Button(self.root, text="Start", bg="#FA8A6F", width=20, command=lambda: self.start())
        self.saveButton=tk.Button(self.root, text="Save current values to file", bg="#FA8A6F", width=20, command=lambda: self.save())
        self.loadButton=tk.Button(self.root, text="Load values from file", bg="white", width=20, command=lambda: self.load())

        self.firstSumLabel = tk.Label(self.root, text="First sum will appear here", bg="#ffffff")
        self.secSumLabel = tk.Label(self.root, text="Second sum will appear here", bg="#ffffff")

        # Entries to get sheets and cols from user
        self.firstSheetEntry = tk.Entry(self.root, width=40)
        self.firstSheetEntry.insert(0, "Enter the first sheet to gather data from")
        self.secSheetEntry = tk.Entry(self.root, width=40)
        self.secSheetEntry.insert(0, "Enter the second sheet to gather data from")
        self.firstColEntry = tk.Entry(self.root, width=40)
        self.firstColEntry.insert(0, "Enter first column to sum here (ex. A)")
        self.secColEntry = tk.Entry(self.root, width=40)
        self.secColEntry.insert(0, "Enter second column to sum here (ex. B)")


        # Set up the grid

        # First Row
        self.wb1Button.grid(row=0, column=0)
        self.firstSheetEntry.grid(row=0, column=1, sticky="NSEW")
        self.firstColEntry.grid(row=0, column=2, sticky="NSEW")

        # Second Row
        self.wb2Button.grid(row=1, column=0)
        self.secSheetEntry.grid(row=1, column=1, sticky="NSEW")
        self.secColEntry.grid(row=1, column=2, sticky="NSEW")

        # Third Row
        self.setButton.grid(row=2, column=0, sticky="NSEW")
        self.startButton.grid(row=2, column=1, sticky="NSEW")

        # Fourth Row
        self.saveButton.grid(row=3, column=0,  sticky="NSEW")
        self.loadButton.grid(row=3, column=1, sticky="NSEW")

        # Info Section
        self.firstSumLabel.grid(row=4, column=0, columnspan=2, sticky="NSW")
        self.secSumLabel.grid(row=4, column=2, columnspan=2, sticky="NSW")
        self.infoLabel.grid(row=5, column=0, columnspan=4, sticky="NSEW")

        self.excelTool = ColumnSummationTool()


    def main(self):

        # Run the GUI Loop. Must execute last
        self.root.mainloop()


    def selectWB(self, workbookNumber):
        fileName = fd.askopenfilename()
        if workbookNumber ==1:
            self.excelTool.openWorkbooks(workbookNumber,fileName)
        elif workbookNumber ==2:
            self.excelTool.openWorkbooks(workbookNumber, fileName)


    def setVals(self):
        sheet1 = str(self.firstSheetEntry.get())
        sheet2 = str(self.secSheetEntry.get())
        col1 = str(self.firstColEntry.get())
        col2 = str(self.secColEntry.get())

        if self.excelTool.wb1 == "" or self.excelTool.wb2 =="":
            self.infoLabel.config(text="Must select excel files first")
            self.infoLabel.config(bg="red")
            self.canSaveToFile = False
            return
        elif sheet1 == "Enter the first sheet to gather data from" or sheet2 == "Enter the second sheet to gather data from":
                self.infoLabel.config(text="Must enter sheet names first")
                self.infoLabel.config(bg="red")
                self.canSaveToFile = False
                return
        elif len(col1) > 5 or len(col2) > 5:
            self.infoLabel.config(text="Invalid column")
            self.infoLabel.config(bg="red")
            self.canSaveToFile = False
            return
        else:
            sheetLoad = self.excelTool.setSheet(sheet1, sheet2)
            if sheetLoad == False:
                self.infoLabel.config(text="Error loading sheets by name (see warnings on console)")
                self.infoLabel.config(bg="red")
                self.canSaveToFile = False
                return
            else:
                self.infoLabel.config(text="Workbook and sheet values set correctly")
                self.canSaveToFile=True
                self.infoLabel.config(bg="#32ed35")
                self.setButton.config(bg="white")
                self.startButton.config(bg="#32ed35")
                self.saveButton.config(bg="#32ed35")



    def save(self):
        if self.canSaveToFile == True:
            wb1 = self.excelTool.wb1Name
            wb2 = self.excelTool.wb2Name
            sheet1 = str(self.firstSheetEntry.get())
            sheet2 = str(self.secSheetEntry.get())
            col1 = str(self.firstColEntry.get())
            col2 = str(self.secColEntry.get())
            saveString = f"{wb1};{wb2};{sheet1};{sheet2};{col1};{col2}"
            with open("config.txt", "w") as f:
                f.write(saveString)
            self.infoLabel.config(text="Config saved")
            self.infoLabel.config(bg="green")
        else:
            self.infoLabel.config(text="Cannot save to config, some values may not be set correctly")
            self.infoLabel.config(bg="red")


    def load(self):

        try:
            with open("config.txt", "r") as f:
                line = f.readline()
            wb1, wb2, sheet1, sheet2, col1, col2 = line.split(";")
            self.excelTool.openWorkbooks(1, wb1)
            self.excelTool.openWorkbooks(2, wb2)
            self.firstSheetEntry.delete(0,tk.END)
            self.firstSheetEntry.insert(0,sheet1)
            self.secSheetEntry.delete(0, tk.END)
            self.secSheetEntry.insert(0, sheet2)
            self.firstColEntry.delete(0, tk.END)
            self.firstColEntry.insert(0, col1)
            self.secColEntry.delete(0, tk.END)
            self.secColEntry.insert(0, col2)
            self.infoLabel.config(text="Config Loaded, click set button, then start button")
            self.infoLabel.config(bg="green")
        except:
            self.infoLabel.config(text="Error loading config, set manually instead")
            self.infoLabel.config(bg="red")











    def start(self):
        col1 = str(self.firstColEntry.get())
        col2 = str(self.secColEntry.get())
        sum1, sum2 = self.excelTool.addColumnValues(col1, col2)
        self.firstSumLabel.config(text=sum1)
        self.secSumLabel.config(text=sum2)
        self.infoLabel.config(text="Values Displayed Above")
        self.infoLabel.config(bg="green")



if __name__ == "__main__":
    c = ColSumWindow()
    c.main()