import win32com.client as win32

LISTA_BC = ("D", "F", "J", "K", "O")

LISTA_AC = ("A", "B", "D", "E" )

class excel():

    def __init__(self):
        self.xl = win32.gencache.EnsureDispatch("Excel.Application")
        self.xl.Visible = True
        self.xl.WindowState = win32.constants.xlMaximized

# ! Funções para a parte do Excel:

    def open_excel(self, path):
        self.xl.Workbooks.Open(path)

    def save_excel(self):
        self.xl.Workbooks.Save()

    def save_as(self, name):
        self.xl.Workbooks.SaveAs(Filename = r''+name+".xlsx")

    def close_excel(self):
        self.xl.Workbooks.Close(True)

    def quit_appclication(self):
        self.xl.Workbooks.Quit()

    def insert_column(self, range, text):
        wrkSht = self.xl.Workbooks.ActiveSheet
        self.rangeObj = wrkSht.Range(range)
        self.rangeObj = wrkSht.Value(text)
        self.rangeObj.EntireColumn.Insert()

    def insert_row(self, range, text):
        wrkSht = self.xlWorkbooks.ActiveSheet
        self.rangeObj = wrkSht.Range(range)
        self.rangeObj = wrkSht.Value(text)
        self.rangeObj.EntireRow.Insert()

    def read_columns(self, i=2, b=0):
        dado = str(self.xl.Worksheets("Sheet1").Range(LISTA_BC[b] + str(i)))

        while dado != None:
            self.xl.Worksheets("Sheet1").Range(LISTA_BC[b] + str(i))
            i+= 1
            b += 1
    
    def read_row(self, i=2):
        CostCenter = str(self.xl.Worksheets("Sheet1").Range(LISTA_AC[0] + str(i)))
        Wbs = str(self.xl.Worksheets("Sheet1").Range(LISTA_AC[1] + str(i)))
        ShortText = str(self.xl.Worksheets("Sheet1").Range(LISTA_AC[2] + str(i)))

        while CostCenter != None:
            CostCenter = str(self.xl.Worksheets("Sheet1").Range(LISTA_AC[0] + str(i)))
            Wbs = str(self.xl.Worksheets("Sheet1").Range(LISTA_AC[1] + str(i)))
            ShortText = str(self.xl.Worksheets("Sheet1").Range(LISTA_AC[2] + str(i)))
            i+= 1
            return CostCenter, Wbs, ShortText