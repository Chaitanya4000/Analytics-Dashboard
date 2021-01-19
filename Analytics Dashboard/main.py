import PySimpleGUI as sg

sg.theme('DarkPurple')

event, values = sg.Window('Get filename example',
                          [[sg.Text('                                     '), sg.Text('TITLE', font=('Arial', 15, 'bold'))],
                              [sg.Text('Filename')], [sg.Input(size=(50,1)), sg.FileBrowse()],
                           [sg.InputCombo(('Priority', 'Language', 'Module', 'Region', 'Day', 'Impact', 'Urgency', 'Risk Assessment', 'Change Scope', 'Description', 'Change Status'), size=(30, 1)),
                            sg.Button('Name', size=(10,1)), sg.Button('Submit', size=(10,1))],
                           [sg.Text('      ')],
                           [sg.Multiline(size=(55,20))]]).read(close=True)

def Buildgraph():
    print("Input"+str(sg.Input))


if event == "Submit":
    Buildgraph()