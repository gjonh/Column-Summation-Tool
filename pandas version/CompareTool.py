import pandas as pd


# This is the second version of this project which works using Pandas
class CompareTool:

    def __init__(self):
        # The first Dataframe
        self.df1 = pd.DataFrame()
        # The second Dataframe
        self.df2 = pd.DataFrame()
        # Initialize the first and second worksheets. These will be selected from a dropdown of
        # Available worksheets later
        self.col1 = ""
        self.col2 = ""


    def openWorkbooks(self, workbookNumber, fileName):
        try:
            if workbookNumber == 1:
                self.df1 = pd.read_csv(fileName)
                print(self.df1)
            elif workbookNumber == 2:
                self.df2 = pd.read_csv(fileName)
                print(self.df2)

            print("Loaded dataframe {num}".format(num=workbookNumber))
        except:
            print("Cannot open the specified file")

    def addColumnValues(self, col1, col2):
        sum1 = 0
        sum2 = 0
        sum1 = self.df1[col1].sum()
        sum2 = self.df2[col2].sum()

        return sum1, sum2
