import PySimpleGUI as sg

sg.theme("DarkGrey4") #The theme of the window for the program
sg.set_options(font=("Arial Bold", 20))
#Menu dropdown from filter button
menu_def = [["filter"], ['Alphabetical', 'Type of Organization']]

buttonmenu = sg.ButtonMenu("Filter", menu_def, size=(50,50))

#Layout of how the window looks
layout = [[sg.Text("Industry Partners List", size=(40, 1), justification="center", expand_x=True)],
          [sg.Menu(menu_def)]
          ]

#Initializing the Window
window = sg.Window("Industry Partners", layout)#Window name, size, and layout

#Loop of executing the window
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED:
        break
window.close()