from typing import ValuesView
import PySimpleGUI as sg
from PySimpleGUI.PySimpleGUI import Checkbox

#integrazione Excel

import pandas as pd


sg.theme('DarkTeal9')   # Colore Tabella 

EXCEL_FILE = 'Database.xlsx'
df = pd.read_excel(EXCEL_FILE)


# Tutto quello all'interno delle finestre 
layout = [  [sg.Text('Nome e Cognome',size=(15,1)), sg.InputText(key= 'Nome e Cognome')],
            [sg.Text('Azienda',size=(15,1)), sg.InputText(key='Azienda')],
            [sg.Text('Mansione',size=(15,1)), sg.InputText(key='Mansione')],
            [sg.Text('Email',size=(15,1)), sg.InputText(key='Email')],
            [sg.Text('Città',size=(15,1)), sg.InputText(key='Città')],
            [sg.Text('Partner',size=(15,1)), sg.InputText(key='Partner')],
            [sg.Text('Azioni svolte',size=(15,1)), sg.InputText(key='Azioni svolte')],
            [sg.Text('Prossimi step',size=(15,1)), sg.InputText(key='Prossimi Step')],
            [sg.Text('Cliente diretto' , size=(15,1)),
                                                sg.Checkbox('Si', key='Si'),
                                                sg.Checkbox('No',   key='No')],
                                                

            [sg.Button('Ok'), sg.Button('Cancella'), sg.Button('Esci')] ]

#se per caso sbaglio e voglio cancellare 'cancella'
def clear_input():
    for key in values:
        window[key]('')
    return None

# Creazione della finestra
window = sg.Window('Compilare i seguenti campi', layout)
# Event Loop to process "events" and get the "values" of the inputs
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Esci': # if user closes window or clicks cancel
        break
    if event == 'Cancella':
        clear_input()
    if event == 'Ok':
        df = df.append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Dati salvati')
        clear_input()

window.close()