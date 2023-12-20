import os
import xlwings as xw
from multiprocessing import Pool

class Excel:
    
    def __init__(self):
        self.wb = None
        self.shape = None
        self.sheet = None
        self.data = None
        self.app = xw.App(visible=False, add_book=False)
        self.app.display_alerts = False
        self.app.screen_updating = False

    def open(self, file_path):
        if os.path.exists(file_path):
            self.wb = self.app.books.open(file_path)
            return True
        else:
            print('file not found')
            return False
    
    def select_sheet(self, index):
        if not index:
            index = 0
        
        if self.wb:
            if isinstance(index, int) or isinstance(index, str):
                self.sheet = self.wb.sheets[index]
            else:
                print('index must be int or str')
        else:
            print('open a workbook first')

    def add_sheet(self, sheet_name):
        if not self.wb:
            self.wb = self.app.books.add()
            self.select_sheet(0)
            self.sheet.name = sheet_name
        else:
            self.sheet = self.wb.sheets.add(sheet_name)

    def cache_data(self, index):
        self.select_sheet(index)
        if self.sheet:
            self.data = self.sheet.used_range.value
            self.shape = [len(self.data), len(self.data[0])]
            del self.sheet
    
    def clear_cache(self):
        self.data = None

    def transpose(self, data):
        return [list(row) for row in zip(*data)]
    
    def write_data(self):
        self.sheet.range((1, 1)).value = self.data
    
    def save(self, file_path):
        self.wb.save(file_path)

    def close(self):
        if self.wb:
            self.wb.close()
        if self.app:
            # self.app.quit()
            self.app.kill()