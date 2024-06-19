import os
import openpyxl
import TkEasyGUI as sg
from generate import generate

layout_data = [
    [
        sg.Text('ファイルの場所'),
        sg.Input(key='data', enable_events=True),
        sg.FileBrowse('Excelファイルを開く', file_types=(('Excelファイル', '*.xlsx'),))
    ],
    [sg.Text('読み込むシート'), sg.Listbox(values=['ファイルを開いてください'], key='sheets')]
]

layout_template = [
    [
        sg.Text('ファイルの場所'),
        sg.Input(key='template', enable_events=True),
        sg.FileBrowse('PowerPointファイルを開く', file_types=(("PowerPointファイル", "*.pptx"),))
    ]
]

layout_save = [
    [
        sg.Text('フォルダーの場所'),
        sg.Input(key='save'),
        sg.FolderBrowse('フォルダー')
    ]
]

layout = [
    [sg.Frame('テンプレート', layout_template)],
    [sg.Frame('名前データ', layout_data)],
    [sg.Frame('保存場所', layout_save)],
    [sg.Button('OK')]
]

window = sg.Window('名札クリエイター', layout)

while window.is_alive():
    event, values = window.read()

    if event == "OK":
        msg = ""
        if not os.path.isfile(values['template']):
            msg += "テンプレートファイルが存在しません\n"
        if not os.path.isfile(values['data']):
            msg += "名前データファイルが存在しません\n"

        if msg != "":
            sg.popup(msg, "エラー")
        else:
            generate(
                values['template'],
                values['data'],
                values['save'],
                values['sheets'][0] if values['sheets'] else None
            )
            break
    elif event == 'data':
        print("DEBUG")
        if values['data'] == '':
            window['sheets'].update(values=["ファイルを開いてください。"])
        elif os.path.isfile(values['data']):
            wb = openpyxl.load_workbook(values['data'])
            window['sheets'].update(values=wb.sheetnames)
            wb.close()
        else:
            window['sheets'].update(values=["ファイルが存在しません。", "開き直してください。"])

window.close()

