import pandas as pd
import csv
from datetime import datetime
import numpy as np
import ctypes
import os


def select_file():
    """Get filepath of folder with html doc from user."""
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Select downloaded HTML file of the printable survey with scores")
    return(file_path)


file = select_file()
all_data = pd.read_html(file)[0]
folder_path = '/'.join(file.split('/')[0:-1])

# pick and assign title to the CSV output
csv_title = str((datetime.now().strftime(
    "%Y.%m.%d_%H.%M%p") + "  Green Globes Export.csv"))


def is_number(s):
    try:
        float(s)
        if float(s) > -100:
            return True
        else:
            return False
    except ValueError:
        return False


all_data2 = pd.DataFrame()
# format dataframe for export
for index, row in all_data.iterrows():
    if is_number(row[1]) and is_number(row[2]) and row[1] != np.NaN and float(row[1]) >= 0:
        #print(index, "yay it's a number:", row[1], type(row[1]))
        all_data2 = all_data2.append(row)
    elif not is_number(row[1]) and not is_number(row[2]) and not is_number(row[3]):
        #print(index, "------its not a number. REMOVE ROW-------------")
        pass
    else:
        #print(index, "------its not a number but there are numbers. Remove value and shift left-------------")
        row[1] = row[2]
        row[2] = row[3]
        row[3] = np.NaN
        all_data2 = all_data2.append(row)


csv_file_path = os.path.join(folder_path, csv_title)
# write dataframe to csv
all_data2.to_csv(csv_file_path, encoding='utf-8')
ctypes.windll.user32.MessageBoxW(
    0, "CSV output file has been saved to: " + csv_file_path, "Complete", 1)
