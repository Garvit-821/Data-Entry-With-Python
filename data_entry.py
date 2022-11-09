from pathlib import Path
import xlsxwriter as xl
import PySimpleGUI as sg
import pandas as pd

# Add some color to the window
sg.theme('DarkBlue7')
#Creating A Worksheet BY Automation
# datasheet = xl.workbook('Data_Entry.xlsv')
# making a worksheet 
# datasheet = datasheet.add_worksheet('first sheet')
# datasheet.write('A1', 'Name'     ) 
# datasheet.write('B1', 'Product'  )
# datasheet.write('C1', 'Age'      )
# datasheet.write('D1', 'Phone no' )
# datasheet.write('E1', 'Email id' )
# datasheet.write('F1', 'City '    )
# datasheet.write('G1', 'Gender'   )
# datasheet.write('H1', 'Tamil'    )
# datasheet.write('I1', 'Hindi'    )
# datasheet.write('J1', 'English'  )
# datasheet.write('K1', 'Children' )





current_dir = Path(__file__).parent if '__file__' in locals() else Path.cwd()
EXCEL_FILE = current_dir / 'Data_Entry.xlsx'
df = pd.read_excel(EXCEL_FILE)

layout = [
    [sg.Text('Made By Garvit Prakash')],
    [sg.Text('Please fill out the following fields:')],


    [sg.Text('Name', size=(15,1)), sg.InputText(key='Name')],
    #name
    [sg.Text(
        'Product', size=(15,1)), sg.InputText(key='Product'
    )],
    [sg.Text(
        'Phone no ', size=(15,1)), sg.InputText(key='Phone No'
        )],
    [sg.Text(
        'Email id', size=(15,1)), sg.InputText(key='Email Id'
        )],
    [sg.Text(
        'City', size=(15,1)), sg.InputText(key='City'
        )],
    [sg.Text(
        'Gender', size=(15,1)), sg.Combo(['Male', 'Female', 'Others'], key='Gender'
        
        )],

        
    [sg.Text('I speak', size=(15,1)),
                            sg.Checkbox('Tamil', key='Tamil'),
                            sg.Checkbox('Hindi', key='Hindi'),
                            sg.Checkbox('English', key='English')],
    [sg.Text('No. of Children', size=(15,1)), sg.Spin([i for i in range(0,16)],
                                                       initial_value=0, key='Children')],
    [sg.Text('Age', size=(15,1)), sg.Spin([i for i in range(0,16)],
                                                       initial_value=0, key='Age')],   




    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
    
]

window = sg.Window('Data Entry Form By Garvit', layout)

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
        new_record = pd.DataFrame(values, index=[0])
        df = pd.concat([df, new_record], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data saved!')
        clear_input()
window.close()
