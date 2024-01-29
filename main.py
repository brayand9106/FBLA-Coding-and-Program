import PySimpleGUI as sg

sg.theme("DarkGrey4") #The theme of the window for the program

#Menu dropdown from filter button
menu_def = ["filter", ['Alphabetical', 'Type of Organization']]

buttonmenu = sg.ButtonMenu("Filter", menu_def, size=(50,50))

#Layout of how the window looks
layout = [[sg.Text("Industry Partners List", size=(100,100), justification="center")]]

#Initializing the Window
window = sg.Window("Industry Partners", layout, size=(800, 400))#Window name, size, and layout

#Loop of executing the window
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED:
        break
window.close()