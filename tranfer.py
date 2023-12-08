import os
import time
import numpy as np
import xlwings as xw

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
            self.data = np.array(self.sheet.used_range.value)
            self.shape = self.sheet.api.Shapes

            del self.sheet
    
    def clear_cache(self):
        self.data = None

    def transpose(self, data):
        if isinstance(data, list):
            data = np.array(data)
        data = np.transpose(data)
        return data
    
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

class Excel_trans:
    def __init__(self):
        self.new = Excel()
        self.old = Excel()
    
    def close(self):
        self.old.close()
        self.new.close()
    
    def trans(self, old_path, new_path):
        self.old.open(old_path)
        self.old.cache_data(0)
        print('self.old.data', type(self.old.data), self.old.data)
        self.new.add_sheet('new')
        self.new.data = self.old.data
        
        # set the bolder
        self.new.sheet.range((1, 1), (self.old.data.shape[0], self.old.data.shape[1])).api.Borders.LineStyle = 1
        
        # set the color, merge and font
        color_str = 'A1:Z1, A2'
        merge_str = ''
        font_str = ''
        self.new.sheet.api.Range(color_str).Interior.ColorIndex  = 3
        self.new.sheet.api.Range(merge_str).Merge()
        self.new.sheet.api.Range(font_str).Font.Bold = True
        self.new.sheet.api.Range(font_str).Font.Size = 14
        self.new.sheet.api.Range(font_str).Font.Name = 'Times New Roman'
        self.new.sheet.api.Range(font_str).HorizontallyAligned = -4108
        
        # set the column width and row length
        self.new.sheet.autofit()
        for i in range(1, self.old.data.shape[1] + 1):
            if self.new.sheet.range(1, i).column_width > 15:
                self.new.sheet.range(1, i).column_width = 15
        for i in range(1, self.old.data.shape[0] + 1):
            if self.new.sheet.range(i, 1).row_height > 15:
                self.new.sheet.range(i, 1).row_height = 15
        
        self.new.write_data()
        self.new.save(new_path)

if __name__ == "__main__":
    mod = input('choose a mod 1 or 2:\n1 for single file\n2 for multiple files in 1 folder')
    
    if mod == '1':
        input_path = input('input_path')
        output_path = input('output_path')
        start_time = time.time()
        trans = Excel_trans()
        trans.trans(input_path, output_path)
        end_time = time.time()
        trans.close()
        print('process complete')
        print(f'cost time: {end_time - start_time}')
    
    elif mod == '2':
        Root_input_path = input('Root_path')
        Root_output_path = input('Root_output_path')
        start_time = time.time()
        for file in os.listdir(Root_input_path):
            trans = Excel_trans()
            trans.trans(f'{Root_input_path}{file}', f'{Root_output_path}\\{file}')
        end_time = time.time()
        trans.close()
        print('process complete')
        print(f'cost time: {end_time - start_time}')
    else:
        print('wrong input')