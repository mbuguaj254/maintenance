#importing pysimplegui for the unterface
import PySimpleGUI as sg
#importing pandas for interacting with excel
import pandas as pd

# Add some color to the window
sg.theme('Sandy Beach')

EXCEL_FILE = 'Data_Entry.xlsx'
df = pd.read_excel(EXCEL_FILE)

layout = [
    [sg.Text('Please fill out the following fields:')], #plain text
    [sg.Text('No.', size=(15,1)), sg.InputText(key='No.')], #input text values
    [sg.Text('User', size=(15,1)), sg.InputText(key='User')], #input text values
    [sg.Text('Division/Department', size=(15,1)), sg.InputText(key='Division/Department')], #input text values
    [sg.Text('Equipment description', size=(15,1)), sg.Combo(['Laptop', 'Desktop', 'IP Phone', 'Monitor', 'Printer'], key='Equipment description')], #drop down options
    [sg.Text('Eqpt. make & model', size=(15,1)), sg.InputText(key='Equipment make & model')],#input text values
    [sg.Text('Serial No.', size=(15,1)), sg.InputText(key='Serial No.')],#input text values
    [sg.Text('GDC Tag No.', size=(15,1)), sg.InputText(key='GDC Tag No.')],    #input text values
    [sg.Text('Hardware Specs.', size=(15,1)), sg.InputText(key='Hardware Specification')],#input text values
    [sg.Text('Software/Apps', size=(15,1)), sg.InputText(key='Software/Applications')],#input text values
    [sg.Text('Status', size=(15,1)), sg.InputText(key='Status')],#input text values
    [sg.Text('Remarks/Comments', size=(15,1)), sg.InputText(key='Technician Remarks/Comments')],    #input text values                    
    [sg.Submit(), sg.Button('Clear'), sg.Exit()] #buttons
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
