import PySimpleGUI as sg

sg.theme("DarkGrey4") #The theme of the window for the program
#Menu dropdown from filter button
menu_def = [['Filter', ['Alphabetical', 'Type of Organization', 'Date']]]

#Button File menu on topleft to make filter button for alphabetical
buttonmenu = sg.ButtonMenu("Filter", menu_def, size=(50,50))


######## TABLE VALUES AND TABLE LAYOUT ##########
partnerCategories = ['Organization', 'Type of Organization', 'Contacts']
partners = [['ASU Ira A. Fulton Schools of Engineering', 'Engineering School', 'FultonSchools@asu.edu'],
            ['Intel Corporation', 'Semi-Conductor Chip Maker', 'intel.partner.marketing.studio@intel.com']

]
partnerTable = sg.Table(values=partners, headings=partnerCategories, font=('Arial', 10), justification= 'center', auto_size_columns=False, max_col_width=50, def_col_width=30, expand_x=True)

#############

##### METHOD FOR CLICKING ORGANIZATIONS ON TABLE########

def getOrganizationFromClick(theEvent):
   # If click is found inside event
   if '+CLICKED+' in theEvent:
       # Getting Row and Column from second tuple in event
       rowAndColumn = theEvent[2]
       

       Organization = rowAndColumn[0]
       return Organization



#######


#Layout of how the window looks
layout = [[sg.Text("Industry Partners List", size=(40, 1), justification="center", expand_x=True, font=("Arial Bold", 20))],
          [sg.Menu(menu_def)], [partnerTable]
          ]

#Initializing the Window
window = sg.Window("Industry Partners", layout, size=(1000,420))#Window name, size, and layout

#Loop of executing the window
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED:
        break
window.close()