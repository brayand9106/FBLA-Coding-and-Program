import PySimpleGUI as sg
import os
from pathlib import Path
import pandas as pd

#Dependencies including openpyxl
sg.theme("DefaultNoMoreNagging") #The theme of the window for the program
#Menu dropdown from filter button
menu_def = [['Filter', ['Alphabetical', 'Type of Organization', 'Date']], ['Help', ['FAQs']]]

#Button File menu on topleft to make filter button for alphabetical
buttonmenu = sg.ButtonMenu("Filter", menu_def, size=(50,50))


######## TABLE VALUES AND TABLE LAYOUT ##########
partnerCategories = ['Organization', 'Type of Organization', 'Contacts']
##### a static list of the partners
global partnersLocked
partnersLocked = [['ASU Ira A. Fulton Schools of Engineering', 'Engineering School', 'FultonSchools@asu.edu'],
                  ['Nvidia Corporation', 'Chip Maker', 'example@email.com'],
            ['Intel Corporation', 'Chip Maker', 'intel.partner.marketing.studio@intel.com']

]
partners = [['ASU Ira A. Fulton Schools of Engineering', 'Engineering School', 'FultonSchools@asu.edu'],
            ['Nvidia Corporation', 'Chip Maker', 'example@email.com'],
            ['Intel Corporation', 'Chip Maker', 'intel.partner.marketing.studio@intel.com']

]
partnerTable = sg.Table(values=partners, headings=partnerCategories, font=('Arial', 10), justification= 'center', auto_size_columns=False, max_col_width=50, def_col_width=30, expand_x=True, key='-TABLE-', enable_click_events=True)

partnerInformation = [[[partnersLocked[0][0]],"Serving and partnering with Faculty, Staff, and Students across the Fulton Schools of Engineering. Our core services include technology planning, support, and implementation."]
                      
                      ]


collectedInformation = []
#############

##### METHOD FOR CLICKING ORGANIZATIONS ON TABLE##########################################################################

def getOrganizationFromClick(theEvent, thePartners):
   # If click is found inside event
   #if theEvent == ('-TABLE-', '+CLICKED+', (tuple)):
       # Getting Row and Column from second tuple in event
    print("Compiling Organization rows and columns...")
    rowAndColumn = theEvent[2]
    print(rowAndColumn)
    if rowAndColumn[0] == -1:
        Organization = 0
    else:
        Organization = rowAndColumn[0]
    print("Organization Compiled from Click!")
       # Getting the organization from the row and column tuple (r, c) from the first element
       #return the organization
    return thePartners[Organization][0]
##########
##### GET POPUP METHOD FOR CLICKING ON ORGANIZATION #######################################################################
def getOrganizationPopup(theOrganization, informationList, userCollection): #User Collection is the list to keep track of partners approved
    print(f"Displaying {theOrganization} Window...")
    theInformation = ""                                                     #informationList is for the information of organizations
    for i in range(len(informationList)):                                  #Getting organization information with theOrganization to use for informationList
        print(f"Traversing List of Information \n Traversing Index {i}") #Debugging
        print(theOrganization) #Debug for our organization
        if theOrganization == informationList[i][0][0]: #if the organization matches with the organization in the info list
            print(informationList[i][0]) #Debug to show information
            theInformation = informationList[i][1] # Packaging all the information from the list of organization
            print(informationList[i][1]) #Debug Confirmation of the Information

#Layout of the popup window for the organization
    layout = [[sg.Text(theOrganization, font=('Arial Bold', 20), justification='center', expand_x=True, size=(20, 1))],
              [sg.Text(theInformation, font=('Arial', 15), expand_x=True, size=(60, 10), auto_size_text=False)],
              [sg.Button('Add Organization', font=("Arial Bold", 8), auto_size_button=False, size=(30, 5), key='AddOrg'), sg.Push(), sg.CButton('Go Back', auto_size_button=False, font=('Arial Bold', 8), size=(30,5))], 
               ]
#Initializing the popup window
    window = sg.Window("Organization Information", layout, size=(800, 420))
#While popup window is active
    print(f"{theOrganization} window is displayed!")    
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

########################################## HELP WINDOW ######################################################
def displayHelpWindow():

    answers = [
        ["To find your saved excel sheet, go to C:/Users/(YOUR USER)/PartnershipBackups.xlsx"]
            ]
    
    layout = [[sg.Text("Frequently Asked Questions:", font=('Arial Bold', 20), justification='center', expand_x=True, size=(20,1))],
              [sg.Button('How do you find saved excel sheet of partners?', font=('Arial Bold', 10), auto_size_button=False, size=(60, 5), key='HELP1')],
               [sg.pin(sg.Text(answers[0][0], justification='left', key='-1-', visible=False, font=('Arial', 12)))],
              [sg.VPush()],
              [sg.CButton('Go Back', auto_size_button=False, font=('Arial Bold', 10), size=(20,5), button_color="Grey")],
              
              ]
    
    window = sg.Window('FAQs', layout, size=(800, 800))

    print("Displaying Help(FAQs) Window!")
    while True:
        event, values = window.read()
        print(event, values)
        if event is not None and event == 'HELP1':
            window['-1-'].Update(visible=True)
        elif event == sg.WIN_CLOSED:
            break
    window.close()
    return partners



########################################## SAVE INFORMATION TO AN EXCEL #####################################
def saveToExcel(allInformation):
    List = pd.DataFrame(columns=partnerCategories, data=allInformation)
    List.to_excel(excel_writer=('PartnershipBackups.xlsx'), sheet_name="Partnered Organizations")
    print(List) 
########################################## UPDATE INFORMATION FROM FILTER #########################
    
def updateInformationFromMenu(filterKeyEvent, theTable):
    #temporary table returned to update table values
    newTable = []
    #If the filter picked is alphabetical:
    if filterKeyEvent == 'Alphabetical':

        #This checks whether first value compared to next value is bigger than other lexographically
        newTable = sorted(theTable, key= lambda x: x[0])
        return newTable
        #Size of partners table to go through
            # for i in range(1, len(theTable)):
            #     print(i)
            #     #This checks whether first value compared to next value is bigger than other lexographically
            #     if theTable[i-1][0] > theTable[i][0]:
            #         if(theTable[i-1] in newTable):
            #             newTable.remove(theTable[i-1])
            #         newTable.insert(i, theTable[i-1])
            #         newTable.insert(i-1, theTable[i])
            #         print("first is bigger")
            #     elif theTable[i-1][0] < theTable[i][0]:
            #         print("second is bigger")
            #         newTable.insert(i-1, theTable[i-1])
            #         newTable.insert(i, theTable[i])
            # return newTable
        #If filter picked is type of organization
    elif filterKeyEvent == 'Type of Organization':
        #This checks whether first value compared to next value is bigger than other lexographically
        newTable = sorted(theTable, key= lambda x: x[1])
        return newTable
        # #Size of partners table to go through
            # for i in range(1, len(theTable)):
            #     print(i)
            #     #This checks whether first value compared to next value is bigger than other lexographically
            #     if theTable[i-1][1] > theTable[i][1]:
            #         newTable.insert(i, theTable[i-1])
            #         newTable.insert(i-1, theTable[i])
            #         print("first is bigger")
            #     elif theTable[i-1][1] < theTable[i][1]:
            #         print("second is bigger")
            #         newTable.insert(i-1, theTable[i-1])
            #         newTable.insert(i, theTable[i])
            #     elif theTable[i-1][1] == theTable[i][1]:
            #         newTable.insert(i-1, theTable[i])
            # print(newTable)
            # return newTable
    elif filterKeyEvent == 'Date':
        #Date feature not implemented yet
        sg.popup_cancel('Not implemented Yet', non_blocking=True,)
        return theTable



    

####### MAIN ########

############################ VIEW INFORMATION WINDOW#############################
def ViewInformationWindow(allInformation):
    layout = [[sg.Text("All Collected Partners", size=(40, 1), justification='center', expand_x=True, font=("Arial Bold", 20))],
              [sg.Table(values= collectedInformation, headings=partnerCategories, font=('Arial', 10), justification= 'center', auto_size_columns=False, max_col_width=50, def_col_width=30, expand_x=True, key='VIEWTABLE', enable_click_events=True)], 
              [sg.Button('Save to Excel', font=("Arial Bold", 8), auto_size_button=False, size=(30, 5), key='SAVTOEXCEL'), sg.Push(), sg.CButton('Go Back', auto_size_button=False, font=('Arial Bold', 8), size=(30,5))], 
              [sg.Button("Remove Org", font=("Arial Bold", 8), auto_size_button=False, size=(30, 5), key=('REMOVEORG'))]]
    
    window = sg.Window("Collected Partners", layout, size=(1000, 420))




### WORK ON THIS #####################
    while True:
        event, values = window.read()
        print(event, values)
        if event == sg.WIN_CLOSED:
            break
        #elif event is not None and ('-VIEWTABLE-' and '+CLICKED+' in event) and (event[2][0] != None):
            #print('viewtableisclicked!')
            #deleteOrg = getOrganizationFromClick(event, collectedInformation)
            #while True:
                #print("in while loop")
                #if event is not None and ('REMOVEORG' in event):
                   #print('executing remove org')
                    #collectedInformation.remove(deleteOrg)
                    #window['-VIEWTABLE-'].Update(values=collectedInformation)
                    #break
        elif event is not None and ('SAVTOEXCEL' in event):
            saveToExcel(collectedInformation)
    window.close()
#################################################################################

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
        getOrganizationPopup((getOrganizationFromClick(event, partners)), partnerInformation, collectedInformation)
    elif event is not None and ('VIEW' in event):
        ViewInformationWindow(collectedInformation)
    elif event is not None and ('Alphabetical' or 'Type of Organization' or 'Date' in event) and (event[2][0] != None):
        if event is not None and event == 'FAQs':
            displayHelpWindow()
            print(partners)
        else:
            window['-TABLE-'].Update(values=(updateInformationFromMenu(event, partners)))
            partners = updateInformationFromMenu(event, partners)
    elif event == sg.WIN_CLOSED:
        break
window.close()