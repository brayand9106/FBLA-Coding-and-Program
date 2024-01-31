import PySimpleGUI as sg
import os

sg.theme("DarkGrey4") #The theme of the window for the program
#Menu dropdown from filter button
menu_def = [['Filter', ['Alphabetical', 'Type of Organization', 'Date']]]

#Button File menu on topleft to make filter button for alphabetical
buttonmenu = sg.ButtonMenu("Filter", menu_def, size=(50,50))


######## TABLE VALUES AND TABLE LAYOUT ##########
partnerCategories = ['Organization', 'Type of Organization', 'Contacts']
##### a static list of the partners
partnersLocked = [['ASU Ira A. Fulton Schools of Engineering', 'Engineering School', 'FultonSchools@asu.edu'],
            ['Intel Corporation', 'Semi-Conductor Chip Maker', 'intel.partner.marketing.studio@intel.com']

]
partners = [['ASU Ira A. Fulton Schools of Engineering', 'Engineering School', 'FultonSchools@asu.edu'],
            ['Intel Corporation', 'Semi-Conductor Chip Maker', 'intel.partner.marketing.studio@intel.com']

]
partnerTable = sg.Table(values=partners, headings=partnerCategories, font=('Arial', 10), justification= 'center', auto_size_columns=False, max_col_width=50, def_col_width=30, expand_x=True, key='-TABLE-', enable_click_events=True)

partnerInformation = [[[partnersLocked[0][0]],"Serving and partnering with Faculty, Staff, and Students across the Fulton Schools of Engineering. Our core services include technology planning, support, and implementation."]
                      
                      ]

#############

##### METHOD FOR CLICKING ORGANIZATIONS ON TABLE########

def getOrganizationFromClick(theEvent, thePartners):
   # If click is found inside event
   #if theEvent == ('-TABLE-', '+CLICKED+', (tuple)):
       # Getting Row and Column from second tuple in event
    print("converting rows and columns")
    rowAndColumn = theEvent[2]
    print(rowAndColumn)
    if rowAndColumn[0] == -1:
        Organization = 0
    else:
        Organization = rowAndColumn[0]
    print(Organization)
       # Getting the organization from the row and column tuple (r, c) from the first element
       #return the organization
    return thePartners[Organization][0]
##########
##### GET POPUP METHOD FOR CLICKING ON ORGANIZATION #######
def getOrganizationPopup(theOrganization, informationList):
    theInformation = ""
    for i in range(len(informationList)):
        print('Hello' + str(i))
        print(theOrganization)
        if theOrganization == informationList[i][0][0]:
            print(informationList[i][0])
            theInformation = informationList[i][1]
            print(informationList[i][1])

    layout = [[sg.Text(theOrganization, font=('Arial Bold', 20), justification='center', expand_x=True, size=(20, 1))],
              [sg.Text(theInformation, font=('Arial', 15), expand_x=True, size=(40, 40), auto_size_text=False)]]

    window = sg.Window("Organization Information", layout, size=(800, 420))
    
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED:
            break
    window.close()



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
    print(event, values)
    if ('-TABLE-' and '+CLICKED+' in event) and (event[2][0] != None):
        getOrganizationPopup((getOrganizationFromClick(event, partnersLocked)), partnerInformation)
    elif event == sg.WIN_CLOSED:
        break
window.close()