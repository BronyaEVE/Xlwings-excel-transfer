# Xlwings-excel-transfer

A tool to transfer one excel file content to another and change the template by using xlwings.  

## How to run

1. Install an **excel** or **wps**.  
2. Install **python3**.  
3. Install **xlwings**.  
4. Prepare the input file path and the output file path.  
5. Run transfer.py.  
6. Choose **1** for **single file**, choose **2** for **multiple files** in 1 folder, choose **3** for **multiprocessing**.

## How to use

1. Make sure you know what you want to do with the data in the old file.
2. All the data in the old file can be found in **self.old.data**, it is a **list** by default, you can use **pd.DataFrame** or **np.array** to replace it.
3. Do all the data transfer in **trans()** method.
4. pass the transfered data to **self.old.data**.
5. set the style you like including border, color, font, merge, column width and row length, etc.

```python
def trans(self, old_path, new_path):
        self.old.open(old_path)
        self.old.cache_data(0)
        self.new.add_sheet('new')
        
        df = self.old.data

        # do your transfer here

        self.new.data = df
        
        # set the bolder
        self.new.sheet.range((1, 1), (len(self.old.data), len(self.old.data[0]))).api.Borders.LineStyle = 1
        
        # set the color, merge and font
        # you should pass a string to the set function, like 'A1:Z1, A2' just like VBA
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
        for i in range(1, len(self.old.data[0]) + 1):
            if self.new.sheet.range(1, i).column_width > 15:
                self.new.sheet.range(1, i).column_width = 15
        for i in range(1, len(self.old.data[0]) + 1):
            if self.new.sheet.range(i, 1).row_height > 15:
                self.new.sheet.range(i, 1).row_height = 15
        
        self.new.write_data()
        self.new.save(new_path)

        if self.old:
            self.old.close()
        if self.new:
            self.new.close()
```

## How to improve performance

1. one read one write.  
2. set color like this.

```python
color_str = 'A1:Z1, A2'
self.new.sheet.api.Range(color_str).Interior.ColorIndex  = 3
```

3. not like this.  

```python

```