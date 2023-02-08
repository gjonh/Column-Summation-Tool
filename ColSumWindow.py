import tkinter as tk
from tkinter import filedialog as fd

class ColSumWindow:

    def __init__(self):
        # This section initializes the tkinter window
        self.root = tk.Tk()
        self.root.title("Summation Verification Tool")
        # WidthxHeight
        self.root.geometry("500x100")
        self.wb1Button = tk.Button(self.root, text="Select Workbook 1", bg="#ffffff",
                              command=lambda: self.selectWB(workbookNumber=1))
        self.wb2Button = tk.Button(self.root, text="Select Workbook 2", bg="#ffffff",
                              command=lambda: self.selectWB(workbookNumber=2))

        # Set up the grid
        self.wb1Button.grid(row=0, column=0)
        self.wb2Button.grid(row=1, column=0)

        # Run the window
        self.root.mainloop()

    def selectWB(self, workbookNumber):
        fileName = fd.askopenfilename()
        if workbookNumber ==1:
            return 1, fileName
            self.root.config
        elif workbookNumber ==2:
            return 2, fileName