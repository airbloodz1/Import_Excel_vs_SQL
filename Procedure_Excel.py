class ProcExec:
    def __init__(self, file_path):
        import win32com.client as win32
        self.excel = win32.gencache.EnsureDispatch('Excel.Application')
        self.excel.Visible = True
        self.workbook = self.excel.Workbooks.Open(file_path)

    def run_macro(self, macro_name):
        self.excel.Application.Run(macro_name)

    def save_and_close(self):
        self.workbook.Save()
        self.workbook.Close()