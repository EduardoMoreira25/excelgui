import PySimpleGUI as sg
import pandas as pd

# Add some color to the window
sg.theme("DarkTeal9")

EXCEL_FILE = "gui-excel.xlsx"
df = pd.read_excel(EXCEL_FILE)

# Each list is a diff row in the gui
layout = [
    [sg.Text("Please fill out the following fields: ")],
    [sg.Text("Name", size=(15,1)), sg.InputText(key="Name")],
    [sg.Text("City", size=(15,1)), sg.InputText(key="City")],
    [sg.Text("Favorite Colour", size=(15,1)), sg.Combo(["Green", "Blue", "Red"], key="Favorite Colour")],
    [sg.Text("I speak", size=(15,1)),
        sg.Checkbox("German", key="German"),
        sg.Checkbox("Portuguese", key="Portuguese"),
        sg.Checkbox("Spanish", key="Spanish")],
    [sg.Text("No. of Children", size=(15,1)), sg.Spin([i for i in range(0,16)],
                                                        initial_value=0, key="Children")],


    [sg.Submit(), sg.Button("Clear"), sg.Exit()]
]

def clear_input():
    for key in values:
        window[key](' ')
    return None

window = sg.Window("Simple data entry form", layout)

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == "Exit":
        break
    if event == "Clear":
        clear_input()
    if event == "Submit":
        df = df.append(values, ignore_index=True)  #keys names and xl columns must match
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup("Data Saved!")
        clear_input()
window.close()


# I LATER BUILT AND EXE TO NOT RUN THE PYTHON SCRIPT ON THE CONSOLE