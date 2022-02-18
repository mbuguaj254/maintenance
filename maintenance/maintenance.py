import PySimpleGUI as sg
import pandas as pd

# Add some color to the window
sg.theme('Sandy Beach')

EXCEL_FILE = 'Data_Entry.xlsx'
df = pd.read_excel(EXCEL_FILE)

layout = [
    [sg.Text('Please fill out the following fields:')],
    [sg.Text('No.', size=(15,1)), sg.InputText(key='No.')],
    [sg.Text('User', size=(15,1)), sg.InputText(key='User')],
    [sg.Text('Division/Department', size=(15,1)), sg.InputText(key='Division/Department')],
    [sg.Text('Equipment description', size=(15,1)), sg.Combo(['Laptop', 'Desktop', 'IP Phone', 'Monitor', 'Printer'], key='Equipment description')],
    [sg.Text('Eqpt. make & model', size=(15,1)), sg.InputText(key='Equipment make & model')],
    [sg.Text('Serial No.', size=(15,1)), sg.InputText(key='Serial No.')],
    [sg.Text('GDC Tag No.', size=(15,1)), sg.InputText(key='GDC Tag No.')],    
    [sg.Text('Hardware Specs.', size=(15,1)), sg.InputText(key='Hardware Specification')],
    [sg.Text('Software/Apps', size=(15,1)), sg.InputText(key='Software/Applications')],
    [sg.Text('Status', size=(15,1)), sg.InputText(key='Status')],
    [sg.Text('Remarks/Comments', size=(15,1)), sg.InputText(key='Technician Remarks/Comments')],                        
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

window = sg.Window('Maintenance sheet', layout)

def clear_input():
    for key in values:
        window[key]('')
    return None

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        clear_input()
    if event == 'Submit':
        df = df.append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data saved!')
        clear_input()
window.close()
