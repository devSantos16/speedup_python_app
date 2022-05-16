import xlsxwriter


class Excel:

    def __init__(self, file_name):
        self.file_name = file_name
        self.workbook = None
        self.worksheet = None

    def init_workbook(self):
        self.workbook = xlsxwriter.Workbook(self.file_name)
        self.worksheet = self.workbook.add_worksheet()

    def end_workbook(self):
        self.workbook.close()

    def write_worksheet(self, celula, response):
        self.worksheet.write(celula, response)








