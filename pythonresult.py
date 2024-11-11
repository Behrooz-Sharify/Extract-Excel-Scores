import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os

window = tk.Tk()
window.title('Extract Data')
window.geometry('400x300')

def upload_file():
    
    file_path = filedialog.askopenfilename(title='Select a file')
    
    if file_path:
        source_file = os.path.basename(file_path)
        sheet_name = 'Class'

        label = tk.Label(window, text=f'{source_file}: is selected')
        label.pack(pady=40)

        df = pd.read_excel(source_file, sheet_name=sheet_name)

        # columns_to_extract = ['Student ID', 'Class ID', 'Time', 'Name', 'Last Name', 'Total_FIT', \
        #  'Total_Windows', 'Total_Word', 'Total_PowerPoint','Total_Excel', 'Total_Internet', 'Kankor',\
        # 'Total Score', 'Percentage', 'Grade']
        
        columns_to_extract = ['Student ID', 'Class ID', 'Time', 'Name', 'Last Name', 'Parent Phone Number',\
            'Total_FIT', 'Total_Windows', ]
        df_selected = df[columns_to_extract]

        output_file = 'output.xlsx'
        df_selected.to_excel(output_file, index=False)

        print('selected columns have been successfully extracted to {source_file}' + '_extracted')

        print(f'File selected: {source_file}')
        #window.destroy()
    else:
        print('No file selected')

upload_button = tk.Button(window, text='Upload a File', command=upload_file)
upload_button.pack(pady=20)

window.mainloop()