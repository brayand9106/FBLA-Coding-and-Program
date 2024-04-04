import PySimpleGUI as sg
import os
from pathlib import Path
import pandas as pd

#Dependencies including openpyxl
sg.theme("DefaultNoMoreNagging") #The theme of the window for the program
#Menu dropdown from filter button
menu_def = [['Filter', ['Alphabetical', 'Type of Organization', 'Show Not Added']], ['Help', ['FAQs']]]

#Button File menu on topleft to make filter button for alphabetical
buttonmenu = sg.ButtonMenu("Filter", menu_def, size=(50,50))


######## TABLE VALUES AND TABLE LAYOUT ##########
partnerCategories = ['Organization', 'Type of Organization', 'Contacts']
##### a static list of the partners
partners = [['ASU Ira A. Fulton Schools of Engineering', 'Engineering School', 'FultonSchools@asu.edu'],
            ['Nvidia Corporation', 'Chip Maker', 'example@email.com'],
            ['Intel Corporation', 'Chip Maker', 'intel.partner.marketing.studio@intel.com'],
            ['Advance Micro Devices', 'Semi-Conductor Chip Makers', 'info@amd.com'], 
            ['Amazon', 'Merchant', 'amazonbusinesscs@amazon.com'], 
            ['Geico', 'Financial Insurance', 'ERSPS@geico.com'], 
            ['BusinessU', 'Business and Finance', 'support@businessu.org'], 
            ['Equidi', 'Education', 'legal@equidi.com.'], 
            ['FICO', 'Business and Finance', 'scoresupport@fico.com'],
            ['Hyatt Hotels', 'Business', 'consumeraffairs@hyatt.com'], 
            ['Juno', 'Business and Finance', 'hello@joinjuno.com'], 
            ['IMA', 'Business and Finance', 'ima@imanet.org'], 
            ['RUBIN', 'Business and Finance', 'chelsea@rubineducation.com'], 
            ['Knowledge Matters', 'Education', 'VBCCentral@KnowledgeMatters.com'],
            ['Eagle University', 'Education', 'taylor@eagleuniversity.org '],
            ['iD Tech', 'Education', 'hello@iDTech.com'],
            ['National Technical Honor Society', 'Education', 'info@nths.org'],
            ['Beta Camp', 'Education', 'hello@beta.camp'],
            ['CareerSafe', 'Business', 'support@careersafeonline.com'],
            ['City Pop', 'Community', 'taylor@citypopdenver.com'],
            ['fund2orgs', 'Community', 'asap@funds2orgs.com'],
            ['March of Dimes', 'Community', 'servicedesk@marchofdimes.org'],
            ['SCAD', 'Education', 'techsupport@scad.edu'],
            ['Meta', 'Education', 'info@meta.edu.np'],
            ['Accountant of International Certified Professional Accountants', 'Education', 'startheregoplaces@aicpa.org']

]

duplicatePartners = [['ASU Ira A. Fulton Schools of Engineering', 'Engineering School', 'FultonSchools@asu.edu'],
            ['Nvidia Corporation', 'Chip Maker', 'example@email.com'],
            ['Intel Corporation', 'Chip Maker', 'intel.partner.marketing.studio@intel.com'],
            ['Advance Micro Devices', 'Semi-Conductor Chip Makers', 'info@amd.com'], 
            ['Amazon', 'Merchant', 'amazonbusinesscs@amazon.com'], 
            ['Geico', 'Financial Insurance', 'ERSPS@geico.com'], 
            ['BusinessU', 'Business and Finance', 'support@businessu.org'], 
            ['Equidi', 'Education', 'legal@equidi.com.'], 
            ['FICO', 'Business and Finance', 'scoresupport@fico.com'],
            ['Hyatt Hotels', 'Business', 'consumeraffairs@hyatt.com'], 
            ['Juno', 'Business and Finance', 'hello@joinjuno.com'], 
            ['IMA', 'Business and Finance', 'ima@imanet.org'], 
            ['RUBIN', 'Business and Finance', 'chelsea@rubineducation.com'], 
            ['Knowledge Matters', 'Education', 'VBCCentral@KnowledgeMatters.com'],
            ['Eagle University', 'Education', 'taylor@eagleuniversity.org '],
            ['iD Tech', 'Education', 'hello@iDTech.com'],
            ['National Technical Honor Society', 'Education', 'info@nths.org'],
            ['Beta Camp', 'Education', 'hello@beta.camp'],
            ['CareerSafe', 'Business', 'support@careersafeonline.com'],
            ['City Pop', 'Community', 'taylor@citypopdenver.com'],
            ['fund2orgs', 'Community', 'asap@funds2orgs.com'],
            ['March of Dimes', 'Community', 'servicedesk@marchofdimes.org'],
            ['SCAD', 'Education', 'techsupport@scad.edu'],
            ['Meta', 'Education', 'info@meta.edu.np'],
            ['Accountant of International Certified Professional Accountants', 'Education', 'startheregoplaces@aicpa.org']

]
partnerTable = sg.Table(values=partners, headings=partnerCategories, font=('Arial', 10), justification= 'center', auto_size_columns=False, max_col_width=50, def_col_width=30, expand_x=True, key='-TABLE-', enable_click_events=True)

partnerInformation = [[[partners[0][0]],"Serving and partnering with Faculty, Staff, and Students across the Fulton Schools of Engineering. Our core services include technology planning, support, and implementation."], 
                      [[partners[23][0]], "Providing the basics of digital marketing to help aid educators with online content such as online quizzes, and lessons through full flexibility"]
                      
                      ]


collectedInformation = []
########################################## SEARCH THROUGH LIST FOR ORGANIZATION #####################################
def searchList(theList, item):
    for i in range(len(theList)):
        if(theList[i][0].find(item) != -1):
            print(theList[i][0].find(item))
            return i


##### METHOD FOR CLICKING ORGANIZATIONS ON TABLE##########################################################################

################### CHECK HERE
def getOrganizationFromClick(theEvent, thePartners):
   # If click is found inside event
   #if theEvent == ('-TABLE-', '+CLICKED+', (tuple)):
       # Getting Row and Column from second tuple in event
    print("Compiling Organization rows and columns...")
    rowAndColumn = theEvent[2]
    print(rowAndColumn)
    if rowAndColumn[0] == -1:
        print("Pick 0th Organization")
        Organization = 0
    else:
        Organization = rowAndColumn[0]
        print("Organization: " + str(Organization))
    print("Organization Compiled from Click!")
       # Getting the organization from the row and column tuple (r, c) from the first element
       #return the organization
    print(thePartners[Organization][0])
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
            for i in range(len(partners)):
                if theOrganization in partners[i] and not(partners[i] in userCollection):
                    userCollection.append(partners[i])
                    sg.popup_ok(partners[i][0] + " has been added!\n Go to \"View Added Information\" to view it!", non_blocking=True)
                    print(userCollection)
        elif event == sg.WIN_CLOSED:
            break
    window.close()

########################################## HELP WINDOW ######################################################
def displayHelpWindow():

    answers = [
        ["To find your saved excel sheet, go to C:/Users/(YOUR USER)/PartnershipBackups.xlsx"],
        ["To remove an organization added, simply go to the \"View Added Information\" and click on the table for any \n organization to get prompted on removing one"],
        ["To fix cell sizes in excel, open the excel sheet file and follow these steps:\n 1. Drag over each cell in excel containing the information\n 2. Make sure excel menu is on \"home\" and look for \"cells\" then click format\n 3. Look for \"Cell Size\" and click \"AutoFit Column Width\""]
            ]
    
    layout = [[sg.Text("Frequently Asked Questions:", font=('Arial Bold', 20), justification='center', expand_x=True, size=(20,1))],
              [sg.Button('How do you find saved excel sheet of partners?', font=('Arial Bold', 10), auto_size_button=False, size=(60, 5), key='HELP1')],
               [sg.pin(sg.Text(answers[0][0], justification='left', key='-1-', visible=False, font=('Arial', 12)))],
               [sg.Button('How do you remove added Organizations?', font=('Arial Bold', 10), auto_size_button=False, size=(60, 5), key='HELP2')],
               [sg.pin(sg.Text(answers[1][0], justification='left', key='-2-', visible=False, font=('Arial', 12)))],
               [sg.Button('How do you fix cell sizes in excel?', font=('Arial Bold', 10), auto_size_button=False, size=(60, 5), key='HELP3')],
               [sg.pin(sg.Text(answers[2][0], justification='left', key='-3-', visible=False, font=('Arial', 12)))],
              [sg.VPush()],
              [sg.CButton('Go Back', auto_size_button=False, font=('Arial Bold', 10), size=(20,5), button_color="Grey")],
              
              ]
    
    window = sg.Window('FAQs', layout, size=(800, 600))

    print("Displaying Help(FAQs) Window!")
    while True:
        event, values = window.read()
        print(event, values)
        if event is not None and event == 'HELP1':
            window['-1-'].Update(visible=True)
        elif event is not None and event == 'HELP2':
            window['-2-'].Update(visible=True)
        elif event is not None and event == 'HELP3':
            window['-3-'].Update(visible=True)
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
    
def updateInformationFromMenu(filterKeyEvent, theTable, acquiredInformation):
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
    elif filterKeyEvent == 'Show Not Added':
        #Date feature not implemented yet
        #sg.popup_cancel('Not implemented Yet', non_blocking=True,)
        #return theTable
        sortedTable = []
        for i in range(len(theTable)):
            if not(theTable[i] in acquiredInformation):
                sortedTable.append(theTable[i])
        return sortedTable



    

####### MAIN ########

############################ VIEW INFORMATION WINDOW#############################
def ViewInformationWindow(allInformation):
    layout = [[sg.Text("All Collected Partners", size=(40, 1), justification='center', expand_x=True, font=("Arial Bold", 20))],
              [sg.Table(values= collectedInformation, headings=partnerCategories, font=('Arial', 10), justification= 'center', auto_size_columns=False, max_col_width=50, def_col_width=30, expand_x=True, key='VIEWTABLE', enable_click_events=True)], 
              [sg.Button('Save to Excel', font=("Arial Bold", 8), auto_size_button=False, size=(30, 5), key='SAVTOEXCEL'), sg.Push(), sg.CButton('Go Back', auto_size_button=False, font=('Arial Bold', 8), size=(30,5))]]
    
    window = sg.Window("Collected Partners", layout, size=(1000, 420))




    while True:
        event, values = window.read()
        print(event, values)
        if event == sg.WIN_CLOSED:
            break
        elif event is not None and ('VIEWTABLE' and '+CLICKED+' in event) and (event[2][0] != None):
            theOrg = getOrganizationFromClick(event, collectedInformation)
            print(theOrg)
            print(collectedInformation)
            the_popup = sg.popup_yes_no("Would you like to remove " + theOrg + "?")
            if(the_popup == "Yes"):
                print(theOrg + " removed!")
                orgIndex = searchList(collectedInformation, theOrg)
                print(orgIndex)
                collectedInformation.pop(orgIndex)
                window['VIEWTABLE'].Update(values=(collectedInformation))
                #collectedInformation.remove(orgIndex)
        elif event is not None and ('SAVTOEXCEL' in event):
            saveToExcel(collectedInformation)
    window.close()
#################################################################################

#Layout of how the window looks
layout = [[sg.Text("Industry Partners List", size=(40, 1), justification="center", expand_x=True, font=("Arial Bold", 20))],
          [sg.Menu(menu_def)], [partnerTable], [sg.Button('View Added Information', auto_size_button=False, font=('Arial Bold', 8), size=(30, 3), key='VIEW'), sg.Push(), sg.CButton("Exit", font=('Arial Bold', 8), auto_size_button=False, size=(30, 3))]
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
        else:
            #This checks for event "Show Not Added" inside of filter
            if event == "Show Not Added":
                #Partners is updated and table with not added organizations
                partners = updateInformationFromMenu(event, partners, collectedInformation)
                window['-TABLE-'].Update(partners)
            #This checks for events other than "Show Not Added" inside of filter
            else:
                #Pulls partners from duplicatePartners to reverse possible deletion through "Show Not Added"
                partners = duplicatePartners
                #Partners is update and table with appropiate filter ("Alphabetical" or "type of organization")
                partners = updateInformationFromMenu(event, partners, collectedInformation)
                window['-TABLE-'].Update(values=(partners))
    elif event == sg.WIN_CLOSED:
        break
window.close()