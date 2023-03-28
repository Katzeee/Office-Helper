import xlwings as xw

class ExcelOperator:
    def __init__(self, file_path, visible=False):
        self.file_path = file_path
        self.wb = None
        self.app = None
        self.visible = visible
    
    def __enter__(self):
        self.app = xw.App(self.visible, add_book=False)
        self.wb = self.app.books.open(self.file_path)
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.wb.save()
        self.wb.close()
        self.app.quit()

    @staticmethod
    def create_excel_file(file_name):
        wb = xw.Book()
        wb.save(file_name)
        wb.close()
    
    def read_cell_value(self, sheet_name, cell_address):
        sheet = self.wb.sheets[sheet_name]
        cell = sheet.range(cell_address)
        return cell.value
    
    def write_cell_value(self, sheet_name, cell_address, value):
        sheet = self.wb.sheets[sheet_name]
        cell = sheet.range(cell_address)
        cell.value = value
    
    def set_cell_color(self, sheet_name, cell_address, color):
        sheet = self.wb.sheets[sheet_name]
        cell = sheet.range(cell_address)
        cell.color = color
    
    def set_cell_font_size(self, sheet_name, cell_address, size):
        sheet = self.wb.sheets[sheet_name]
        cell = sheet.range(cell_address)
        cell.api.Font.Size = size
    
    def set_column_width(self, sheet_name, column, width):
        sheet = self.wb.sheets[sheet_name]
        sheet.range(f"{column}1:{column}{sheet.cells.last_cell.row}").api.ColumnWidth = width

    def save_excel(self):
        self.wb.save()
