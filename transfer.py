import os
import time
from multiprocessing import Pool
from tools import Excel

def worker(transfer, input_path, output_path):
    try:
        transfer.trans(input_path, output_path)
    except Exception as e:
        print(f'Error in worker: {e}')

class Excel_trans:
    def __init__(self):
        self.new = Excel()
        self.old = Excel()
    
    def trans(self, old_path, new_path):
        self.old.open(old_path)
        self.old.cache_data(0)
        self.new.add_sheet('new')
        self.new.data = self.old.data
        self.new.shape = [len(self.new.data), len(self.new.data[0])]
        
        # set the bolder
        self.new.sheet.range((1, 1), (self.old.shape[0], self.old.shape[1])).api.Borders.LineStyle = 1
        
        # set the color, merge and font
        color_str = 'A1:Z1, A2'
        merge_str = 'A1:Z1, A2'
        font_str = 'A1:Z1, A2'
        self.new.sheet.api.Range(color_str).Interior.ColorIndex  = 3
        self.new.sheet.api.Range(merge_str).Merge()
        self.new.sheet.api.Range(font_str).Font.Bold = True
        self.new.sheet.api.Range(font_str).Font.Size = 14
        self.new.sheet.api.Range(font_str).Font.Name = 'Times New Roman'                       
        self.new.sheet.api.Range('A3:Z3').HorizontalAlignment = -4108
        
        # set the column width and row length
        self.new.sheet.autofit()
        for i in range(1, self.new.shape[1] + 1):
            if self.new.sheet.range(1, i).column_width > 15:
                self.new.sheet.range(1, i).column_width = 15
        for i in range(1, self.new.shape[1] + 1):
            if self.new.sheet.range(i, 1).row_height > 15:
                self.new.sheet.range(i, 1).row_height = 15
        
        self.new.write_data()
        self.new.save(new_path)

        if self.old:
            self.old.close()
        if self.new:
            self.new.close()

if __name__ == '__main__':
   try:
       mod = int(input('choose a mod 1 or 2 or 3:\n1 for single file\n2 for multiple files in 1 folder\n3 for multiple process\n'))
       if mod == 1:
           input_path = input('input_path\n')
           output_path = input('output_path\n')
           start_time = time.time()
           trans = Excel_trans()
           trans.trans(input_path, output_path)
           end_time = time.time()
           print('process complete')
           print(f'cost time: {end_time - start_time}')
       
       elif mod == 2:
           Root_input_path = input('Root_input_path\n')
           Root_output_path = input('Root_output_path\n')
           start_time = time.time()
           for file in os.listdir(Root_input_path):
               trans = Excel_trans()
               print('path', f'{Root_input_path}\\{file}', f'{Root_output_path}\\{file}')
               trans.trans(f'{Root_input_path}\\{file}', f'{Root_output_path}\\{file}')
           end_time = time.time()
           print('process complete')
           print(f'cost time: {end_time - start_time}')
       
       elif mod == 3:
           Root_input_path = input('Root_input_path\n')
           Root_output_path = input('Root_output_path\n')
           p = Pool(4)
           start_time = time.time()
           file_names = os.listdir(Root_input_path)
           for i in range(4):
               transfer = Excel_trans()
               print(f'{Root_input_path}\\{file_names[i]}', f'{Root_output_path}\\{file_names[i]}')
               p.apply_async(func=worker(transfer, f'{Root_input_path}\\{file_names[i]}', f'{Root_output_path}\\{file_names[i]}'))
           p.close()
           p.join()
           end_time = time.time()
           print('All processes finished.')
           print(f'cost time: {end_time - start_time}')

       else:
           print('wrong input')

   except Exception as e:
       print(f'Error in main: {e}')