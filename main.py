import PySimpleGUI as sg
import os
from pathlib import Path
import pandas as pd

sg.theme("DarkGrey4") #The theme of the window for the program
#Menu dropdown from filter button
menu_def = [['Filter', ['Alpha betical', 'Type of Organization', 'Date']]]

#Button File menu on topleft to make filter button for alphabetical
buttonmenu = sg.ButtonMenu("Filter", menu_def, size=(50,50))


######## TABLE VALUES AND TABLE LAYOUT ##########
partnerCategories = ['Organization', 'Type of Organization', 'Contacts']
##### a static list of the partners
global partnersLocked
partnersLocked = [['ASU Ira A. Fulton Schools of Engineering', 'Engineering School', 'FultonSchools@asu.edu'],
            ['Intel Corporation', 'Semi-Conductor Chip Maker', 'intel.partner.marketing.studio@intel.com']

]
partners = [['ASU Ira A. Fulton Schools of Engineering', 'Engineering School', 'FultonSchools@asu.edu'],
            ['Intel Corporation', 'Semi-Conductor Chip Maker', 'intel.partner.marketing.studio@intel.com']

]
partnerTable = sg.Table(values=partners, headings=partnerCategories, font=('Arial', 10), justification= 'center', auto_size_columns=False, max_col_width=50, def_col_width=30, expand_x=True, key='-TABLE-', enable_click_events=True)

partnerInformation = [[[partnersLocked[0][0]],"Serving and partnering with Faculty, Staff, and Students across the Fulton Schools of Engineering. Our core services include technology planning, support, and implementation."]
                      
                      ]


collectedInformation = []
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
def getOrganizationPopup(theOrganization, informationList, userCollection): #User Collection is the list to keep track of partners approved
    theInformation = ""                                                     #informationList is for the information of organizations
    for i in range(len(informationList)):                                  #Getting organization information with theOrganization to use for informationList
        print('Hello' + str(i)) #Debugging
        print(theOrganization) #Debug for our organization
        if theOrganization == informationList[i][0][0]: #if the organization matches with the organization in the info list
            print(informationList[i][0]) #Debug to show information
            theInformation = informationList[i][1] # Packaging all the information from the list of organization
            print(informationList[i][1]) #Debug Confirmation of the Information

#Layout of the popup window for the organization
    layout = [[sg.Text(theOrganization, font=('Arial Bold', 20), justification='center', expand_x=True, size=(20, 1))],
              [sg.Text(theInformation, font=('Arial', 15), expand_x=True, size=(60, 10), auto_size_text=False)],
              [sg.Button('Add Organization', font=("Arial Bold", 8), auto_size_button=False, size=(30, 5), key='AddOrg'), sg.Push(), sg.CButton('Close', auto_size_button=False, font=('Arial Bold', 8), size=(30,5))], 
               ]
#Initializing the popup window
    window = sg.Window("Organization Information", layout, size=(800, 420))

#While popup window is active    
    while True:
        event, values = window.read()
        print(event)
        if event is not None and 'AddOrg' in event: #Check if "Add Organization button is clicked to append"
            for i in range(len(partnersLocked)):
                if theOrganization in partnersLocked[i] and not(partnersLocked[i] in userCollection):
                    userCollection.append(partnersLocked[i])
                    print(userCollection)
        elif event == sg.WIN_CLOSED:
            break
    window.close()

############################ VIEW INFORMATION WINDOW#############################
def ViewInformationWindow(allInformation):
    layout = [[sg.Text("All Collected Partners", size=(40, 1), justification='center', expand_x=True, font=("Arial Bold", 20))],
              [sg.Table(values= collectedInformation, headings=partnerCategories, font=('Arial', 10), justification= 'center', auto_size_columns=False, max_col_width=50, def_col_width=30, expand_x=True, key='VIEWTABLE', enable_click_events=True)], 
              [sg.Button('Save to Excel', font=("Arial Bold", 8), auto_size_button=False, size=(30, 5), key='SAVTOEXCEL'), sg.Push(), sg.CButton('Close', auto_size_button=False, font=('Arial Bold', 8), size=(30,5))]]
    
    window = sg.Window("Collected Partners", layout, size=(1000, 420))

    while True:
        event, values = window.read()
        print(event, values)
        if event == sg.WIN_CLOSED:
            break
    window.close()
#################################################################################


#WORK ON THIS
########################################## SAVE INFORMATION TO AN EXCEL #####################################
# def saveToExcel(allInformation):
    #p d.DataFrame.to_excel(excel_writer=(str(Path.home() / "Downloads")), sheet_name="Partnered Organizations"  ) 


####### MAIN ########


#Layout of how the window looks
layout = [[sg.Text("Industry Partners List", size=(40, 1), justification="center", expand_x=True, font=("Arial Bold", 20))],
          [sg.Menu(menu_def)], [partnerTable], [sg.Button('View Added Information', auto_size_button=False, font=('Arial Bold', 8), size=(30, 3), key='VIEW')]
          ]

#Initializing the Window
window = sg.Window("Industry Partners", layout, size=(1000,420))#Window name, size, and layout

#Loop of executing the window
while True:
    event, values = window.read()
    print(event, values)
    if event is not None and ('-TABLE-' and '+CLICKED+' in event) and (event[2][0] != None):
        getOrganizationPopup((getOrganizationFromClick(event, partnersLocked)), partnerInformation, collectedInformation)
    elif event is not None and ('VIEW' in event):
        ViewInformationWindow(collectedInformation)
    elif event == sg.WIN_CLOSED:
        break
window.close()