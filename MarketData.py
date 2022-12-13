# Raw Package
import pandas as pd
#import the API and set it to a variable that will be able to call the methods of the API
import yfinance as yf
#GUI Source
import PySimpleGUI as sg


sg.theme('Reddit')
sg.set_options(font='calibri', text_color='black')

def convert_to_excel(data, output_folder, fileName):
    df = pd.DataFrame.from_dict(dict(data))
    df.to_excel(output_folder+'/'+fileName+'.xlsx')
    
layout = [
    [sg.Text("Information entered below:")],
    [sg.Text("Ticker:"),sg.InputText(key = 't'),sg.Push()],
    [sg.Text("Period:"), sg.InputText(key = 'p'), sg.Push()],
    [sg.Text("Output:"), sg.InputText(key = '-folder-'), sg.FolderBrowse(), sg.Push()],
    [sg.Text("Save As:"), sg.InputText(key = '-filename-'), sg.Push()],
    [sg.Exit(), sg.Button("OK")]
]

window = sg.Window("Market Data", layout)

while True:
    try:
        event, values = window.read()
        if event == "OK":
            data = yf.download(tickers = values['t'], period = values['p'], interval = '1d')
            convert_to_excel(data, values['-folder-'], values['-filename-'])
            sg.popup_no_titlebar("Complete!")

        if event == sg.WIN_CLOSED or event == "Exit":
            window.close()
            break
    
    except:
        sg.popup_no_titlebar("error")

