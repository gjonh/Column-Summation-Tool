import tkinter as tk
from tkinter import filedialog as fd
from ColumnSummationTool import ColumnSummationTool

class ColSumWindow:

    def __init__(self):
        # This section initializes the tkinter window
        self.root = tk.Tk()
        self.root.title("Summation Verification Tool")
        # WidthxHeight
        self.root.geometry("800x120")
        self.wb1Button = tk.Button(self.root, width=20, text="Select Workbook 1", bg="#ffffff",
                              command=lambda: self.selectWB(workbookNumber=1))
        self.wb2Button = tk.Button(self.root, width=20, text="Select Workbook 2", bg="#ffffff",
                              command=lambda: self.selectWB(workbookNumber=2))
        self.infoLabel = tk.Label(self.root, text="Set both sheets and columns before starting", bg="#FA8A6F")
        self.setButton=tk.Button(self.root, text="Set Values", bg="#32ed35", width=20, command=lambda: self.setVals())
        self.startButton=tk.Button(self.root, text="Start", bg="#FA8A6F", width=20, command=lambda: self.start())
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

        # Info Section
        self.firstSumLabel.grid(row=3, column=0, columnspan=2, sticky="NSW")
        self.secSumLabel.grid(row=3, column=2, columnspan=2, sticky="NSW")
        self.infoLabel.grid(row=4, column=0, columnspan=4, sticky="NSEW")

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

    def start(self):
        sheet1= str(self.firstSheetEntry.get())
        sheet2 = str(self.secSheetEntry.get())


if __name__ == "__main__":
    c = ColSumWindow()
    c.main()