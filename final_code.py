import pandas as pd
import os
from tkinter import *
import logging
import time

contract = []
valid_from = []
valid_to = []
port_of_loading = []
port_of_discharge = []
_20DC = []
_40DC = []
_40HC = []
currency = []
sheet = []

contract_is_valid = True

def work(df, name):
    start_index = 0
    end_index = 0

    for i in range(df.shape[0]):
        if df.iloc[i, 0] == 'AREA':
            start_index = i
            break

    for i in range(start_index + 1, df.shape[0]):
        if pd.isna(df.iloc[i, 0]):
            end_index = i
            break
        contract_name = df.iloc[2, 2]
        try:
            assert contract_name == 'QHOF113592'
        except AssertionError:
            global contract_is_valid
            contract_is_valid = False

        contract.append(contract_name)
        valid_from.append(pd.to_datetime(df.iloc[4, 2]).date())
        valid_to.append(pd.to_datetime(df.iloc[4, 2]).date())
        currency.append('USD')
        sheet.append(name)

    df = df.iloc[start_index:end_index]
    df.columns = df.iloc[0]
    df = df.iloc[1:]

    df.reset_index(inplace=True)

    for i in range(df.shape[0]):
        port_of_loading.append(df.loc[i, 'PORT OF LOADING'])
        port_of_discharge.append(df.loc[i, 'PORT OF DISCHARGE'])
        _20DC.append(df.loc[i, '20\' usd'])
        _40DC.append(df.loc[i, '40\' usd'])
        _40HC.append(df.loc[i, 'HC usd'])

username = os.getlogin()
file = pd.ExcelFile(f'C:\\Users\\{username}\\Desktop\\DP_project_file.xlsx')

for sheet_name in file.sheet_names:
    if not 'TAO' in sheet_name:
        work(pd.read_excel(file, sheet_name), sheet_name)

data = {'contract': contract,
        'valid_from': valid_from,
        'valid_to': valid_to,
        'port_of_loading': port_of_loading,
        'port_of_discharge': port_of_discharge,
        '20DC': _20DC,
        '40DC': _40DC,
        '40HC': _40HC,
        'currency': currency,
        'sheet': sheet}

df_new = pd.DataFrame(data)

root = Tk()

path = f'C:\\Users\\{username}\\Desktop\\DP_Project_Results'
if not os.path.isdir(path):
    os.mkdir(path)
completeName = os.path.join(path, 'Result.xlsx')
df_new.to_excel(completeName, index=False)

logger=logging.getLogger(__name__)
file_handler=logging.FileHandler(f'{path}\\Log.log')
logger.addHandler(file_handler)
logger.setLevel(logging.INFO)
formatter=logging.Formatter('%(asctime)s:%(levelname)s:%(name)s:%(message)s')
file_handler.setFormatter(formatter)
created=str(time.ctime(os.path.getctime(path)))
logger.info(f'the file located at the path {path} was created at {created}')

options = Menu(root)
root.config(menu=options)
sub= Menu(options, tearoff=0)

def open_file():
    os.startfile(completeName)

def open_folder():
    os.system(f'start {path}')

def leave():
    root.destroy()

options.add_cascade(label='Menu', menu=sub)
sub.add_command(label='Open the result file', command=open_file)
sub.add_command(label='Open the new folder', command=open_folder)

label = Label(root, text='Your file was successfully saved. If you want to open the result file or the new folder, select the menu')
label.grid(row=0, columnspan=3)
button = Button(root, text='Exit', command=leave)
button.grid(row=2, column=3)
root.mainloop()

if not contract_is_valid:
    print("The contract is different. Double check if the file was processed correctly")
