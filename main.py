import PySimpleGUI as sg
import pandas as pd

#Dependencies including openpyxl
sg.theme("DefaultNoMoreNagging") #The theme of the window for the program
#Menu dropdown from filter button
menu_def = [['Filter', ['Alphabetical', 'Type of Organization', 'Show Not Added']], ['Help', ['FAQs']]]

#Button File menu on topleft to make filter button for alphabetical
buttonmenu = sg.ButtonMenu("Filter", menu_def, size=(50,50))


########################################## TABLE VALUES AND TABLE LAYOUT #####################################



partnerCategories = ['Organization', 'Type of Organization', 'Contacts']

#Mutable table to be used during display on program for organizations
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
            ['AICPA', 'Education', 'startheregoplaces@aicpa.org']

]

#Immutable table to be used to recover lost data from table
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
            ['AICPA', 'Education', 'startheregoplaces@aicpa.org']

]
partnerTable = sg.Table(values=partners, headings=partnerCategories, font=('Arial', 10), justification= 'center', auto_size_columns=False, max_col_width=50, def_col_width=30, expand_x=True, key='-TABLE-', enable_click_events=True)

#Information for all the partners
partnerInformation = [[[partners[0][0]],"Serving and partnering with Faculty, Staff, and Students across the Fulton Schools of Engineering. Our core services include technology planning, support, and implementation."], 
                     [[partners[1][0]], "A semiconductor company focusing exclusively on the production of Graphics Processing Units."],
                      [[partners[2][0]], "Intel corporation both designs and manufactures a variety of computer parts and other such related products."],
                      [[partners[3][0]], "Microchip company that develops a variety of computer parts for businesses and consumer markets."],
                      [[partners[4][0]], "Every item purchased on Amazon supports FBLA."],
                      [[partners[5][0]], "Providing FBLA members with car insurance at a special discount."],
                      [[partners[6][0]], "Preparing Students to become proficcient in the business field with courses including Marketing, Accounting, Entrepreneurship, and much more."],
                      [[partners[7][0]], "A platform to showcase the many skills that students have that allows for the connection between other students in a professional manner."],
                      [[partners[8][0]], "Educating students on the importance of credit literacy through free education in fun and simple ways for the most effective learning experience."],
                      [[partners[9][0]], "Providing FBLA with hotels for various events and conferences for competitions and more."],
                      [[partners[10][0]], "Starting from Harvard Business School, Juno is made to alleviate the pressure put on students by student loans by negotiating savings for free."],
                      [[partners[11][0]], "IMA offers finance and business students with the necesary information for success through conferences, scholarships, competitions and much more."],
                      [[partners[12][0]], "Teaching students how to how to improve employability skills by teaching proper ways to write emails."],
                      [[partners[13][0]], "Compete in a competitive event that test the skills needed for future finances such as opening bank accounts, saving, and budgeting"],
                      [[partners[14][0]], "Giving students a 7 year headstart on their career, Eagle University is open to students 15-25 years old with career planning, performmance strategy, and much more."],
                      [[partners[15][0]], "Preparing students for college by enhancing various skills, iD Tech has everything covered ranging from private lessons, to Summer camps to ensure the best outcome for college admissions."],
                      [[partners[16][0]], "Acknowledging the hard work of CTE students, and FBLA memebers alike, NTHS encourages students to perform at their best, and eventually work in a skilled workspace with the many scholarships given."],
                      [[partners[17][0]], "A virtual program for high school students to make more future entrepreneurs and leaders, featuring a fully interactive online workspace to ensure the proper training of students."],
                      [[partners[18][0]], "Training more than 1.2 million students, CarrerSafe focuses on providing students with a safe future by giving them proper training for online safety."],
                      [[partners[19][0]], "A new initiative way to do fundraisers, providing FBLA chapters with product lines that gives half of the money received back with no hassles or difficulties."],
                      [[partners[20][0]], "Create fundraising goals for all FBLA chapters with tools provided for by fund2org, helping not only FBLA, but those in need."],
                      [[partners[21][0]], "March of Dimes aims to help aid mothers and babies through the advocation of health by raising money to assist soon to be mothers, and mothers in need."],
                      [[partners[22][0]], "Bringing students from more than 100 countries with over 100 academic degree programs. Internships, certifications, and collaborative projects are all involved in the SCAD curriculum."],
                      [[partners[23][0]], "Providing the basics of digital marketing to help aid educators with online content such as online quizzes, and lessons through full flexibility"],
                      [[partners[24][0]], "(Accountant of International Certified Professional Accountants): \nLearn how to become a certified professional accountant with a fully personalized College Checklist, as well as the many opportunities to learn from professionals."]
                      
                      ]

#Organizations added by user
collectedInformation = []



########################################## SEARCH THROUGH LIST FOR ORGANIZATION #####################################


#This searches through a list to find a specific organization returning the index of it
def searchList(theList, item):
    for i in range(len(theList)):
        if(theList[i][0].find(item) != -1):
            print(theList[i][0].find(item))
            return i



########################################## METHOD FOR CLICKING ORGANIZATIONS ON A TABLE #####################################



def getOrganizationFromClick(theEvent, thePartners):
    # If click is found inside event
    # Getting Row and Column from second tuple in event
    print("Compiling Organization rows and columns...")
    rowAndColumn = theEvent[2]
    print("Selected cell is " + str(rowAndColumn))
    if rowAndColumn[0] == -1:
        print("Override selection of headings by selecting first avaiable organization.")
        Organization = 0
    else:
        Organization = rowAndColumn[0]
        print("Selected Organization postition in partners is " + str(Organization))
    print("Organization Compiled from Click!")
       # Getting the organization from the row and column tuple (r, c) from the first element
       #return the organization
    print("The first selected partner is " + thePartners[Organization][0])
    return thePartners[Organization][0]



########################################## GET POPUP INFORMATION FOR CLICKING ON ORGANIZATION #####################################



def getOrganizationPopup(theOrganization, informationList, userCollection, theToggle): #User Collection is the list to keep track of partners approved
    print(f"Displaying {theOrganization} Window...")
    theInformation = ""                                                     #informationList is for the information of organizations
    for i in range(len(informationList)):                                  #Getting organization information with theOrganization to use for informationList
        print(f"Traversing List of Information \n Traversing Index {i}") #Debugging
        print("The selected Organization is " + str(theOrganization)) #Debug for our organization
        if theOrganization == informationList[i][0][0]: #if the organization matches with the organization in the info list
            print("The selected organization is " + str(informationList[i][0]) + " compiled information about them is: " + str(informationList[i][1])) #Debug to confirm information
            theInformation = informationList[i][1] # Packaging all the information from the list of organization

#Layout of the popup window for the organization
    layout = [[sg.Text(theOrganization, font=('Arial Bold', 20), justification='center', expand_x=True, size=(20, 1))],
              [sg.Text(theInformation, font=('Arial', 15), expand_x=True, size=(60, 10), auto_size_text=False)],
              [sg.Button('Add Organization', font=("Arial Bold", 8), auto_size_button=False, size=(30, 5), key='AddOrg'), sg.Push(), sg.CButton('Go Back', auto_size_button=False, font=('Arial Bold', 8), size=(30,5))], 
               ]
    
#Initializing the popup window
    window = sg.Window("Organization Information", layout, size=(800, 420))
#While popup window is active
    print(f"{theOrganization} window is being displayed!")   
    while True:
        event, values = window.read()
        print("The user has selected: " + str(event))
        if event is not None and 'AddOrg' in event: #Check if "Add Organization button is clicked to append"
            for i in range(len(partners)):
                if theOrganization in partners[i] and not(partners[i] in userCollection):
                    userCollection.append(partners[i])
                    #Checks if dynamicToggle is enabled to save to excel if organization appended
                    if theToggle == True:
                        saveToExcel(userCollection, theToggle)
                    sg.popup_ok(partners[i][0] + " has been added!\n Go to \"View Added Information\" to view it!", non_blocking=True)
                    print("The user selected organizations are: " + str(userCollection))
        elif event == sg.WIN_CLOSED:
            break
    window.close()



########################################## HELP WINDOW #####################################
    


def displayHelpWindow():

    #Answers for each FAQs in the help menu
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
    #Displays the help answer in the FAQ based on user response
    while True:
        event, values = window.read()
        print("The user has selected " + str(event) + " in the faq window.")
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


#Converts the list to a pandas dataframe(special table) to an excel file
def saveToExcel(allInformation, theToggle=False):
    #Converting list into a pandas dataframe
    List = pd.DataFrame(columns=partnerCategories, data=allInformation)
    #Pandas dataframe into excel
    List.to_excel(excel_writer=('PartnershipBackups.xlsx'), sheet_name="Partnered Organizations")
    if theToggle == False:
        sg.popup_ok("A backup with your added organizations has been made called \"PartnershipBackups.xlsx\" in your user file!", non_blocking=True) 



########################################## UPDATE INFORMATION FROM FILTER ###################################
    


def updateInformationFromMenu(filterKeyEvent, theTable, acquiredInformation):
    #temporary table returned to update table values
    newTable = []
    #If the filter picked is alphabetical:
    if filterKeyEvent == 'Alphabetical':

        #This checks whether first value compared to next value is bigger than other lexographically
        newTable = sorted(theTable, key= lambda x: x[0])
        return newTable

    #If filter picked is type of organization
    elif filterKeyEvent == 'Type of Organization':
        #This checks whether first value compared to next value is bigger than other lexographically
        newTable = sorted(theTable, key= lambda x: x[1])
        return newTable
    
    #If filter picked is 'show not added'
    elif filterKeyEvent == 'Show Not Added':
        #Sorts partner table to be put in a temporary sorted table that is returned showing not added organizations
        sortedTable = []
        for i in range(len(theTable)):
            if not(theTable[i] in acquiredInformation):
                sortedTable.append(theTable[i])
        return sortedTable




############################ VIEW INFORMATION WINDOW #############################
    


def ViewInformationWindow(allInformation, theToggle):
    layout = [[sg.Text("All Collected Partners", size=(40, 1), justification='center', expand_x=True, font=("Arial Bold", 20))],
              [sg.Table(values= collectedInformation, headings=partnerCategories, font=('Arial', 10), justification= 'center', auto_size_columns=False, max_col_width=50, def_col_width=30, expand_x=True, key='VIEWTABLE', enable_click_events=True)], 
              [sg.Button('Save to Excel', font=("Arial Bold", 8), auto_size_button=False, size=(30, 5), key='SAVTOEXCEL'), sg.Push(), sg.CButton('Go Back', auto_size_button=False, font=('Arial Bold', 8), size=(30,5))],
              [sg.Button('Restore from Backup', font=("Arial Bold", 8), auto_size_button=False, size=(30, 5), key='-RESTORE-')]]
    
    window = sg.Window("Collected Partners", layout, size=(1000, 420))

    while True:
        event, values = window.read()
        print("The user has selected " + str(event) + " with " + str(values) + " in the information window.")
        if event == sg.WIN_CLOSED:
            break
        
        #This displays the option from organization clicked on to be removed
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
                #Checks if dynamicToggle is enabled to save removed changes
                if theToggle == True:
                    saveToExcel(collectedInformation, theToggle)
                window['VIEWTABLE'].Update(values=(collectedInformation))
        #Saves collected information into excel
        elif event is not None and ('SAVTOEXCEL' in event):
            saveToExcel(collectedInformation)

        #Restores backup into collected information to edit
        elif event is not None and ('-RESTORE-' in event):
            try:
                #Opens the backup file
                theBackup = pd.read_excel("PartnershipBackups.xlsx")
                #This loops through each row of the pandas dataframe
                for row in theBackup.itertuples():
                    #This creates a temporary table to append each information of organization to eventually append into
                    #collectedInformation as a whole
                    elementOrganization = []
                    for i in range(2, 5):
                        elementOrganization.append(str(row[i]))
                    collectedInformation.append(elementOrganization)
                    #The view table is updated with restored backups
                    window['VIEWTABLE'].Update(values=collectedInformation)
                sg.popup_ok("Partnership Backups Restored!")
            except FileNotFoundError:
                print("No file \"PartnershipBackups.xlsx\" was found, either deleted/does not exist or name changed")
                sg.popup_ok("Could not find \"PartnershipBackups.xlsx\", either deleted/does not exist or name changed")

    window.close()



########################################## MAIN PRIMARY WINDOW #####################################



#Layout of how the window looks
layout = [[sg.Text("Industry Partners List", size=(40, 1), justification="center", expand_x=True, font=("Arial Bold", 20))],
          [sg.Menu(menu_def)], [partnerTable], [sg.Button('View Added Information', auto_size_button=False, font=('Arial Bold', 8), size=(30, 3), key='VIEW'), sg.Push(), sg.CButton("Exit", font=('Arial Bold', 8), auto_size_button=False, size=(30, 3))],
          [sg.Push(), sg.Text("Dynamic Backup Toggle", font=('Arial', 10), size=(15, 2)), sg.Button("Off", font=('Arial Bold', 8), size=(15,2), button_color="white on red", key='TOGGLE')]
          ]

#Initializing the Window
window = sg.Window("Industry Partners", layout, size=(1000,420))#Window name, size, and layout

#Checks if dynamic saving is toggled
dynamicToggle = False

#Loop of executing the primary window
while True:
    event, values = window.read()
    print("The user has selected " + str(event) + " with " + str(values) + " in the primary window.")

    #Checks if dynamicToggle button is activated
    if event is not None and ('TOGGLE' in event):
        #If button is false then set it to true
        if dynamicToggle == False:
            dynamicToggle = True
            window['TOGGLE'].Update(text='On', button_color='white on green')
            sg.popup_ok("Dynamic toggle is On! Any organizations added/removed will be automatically saved to excel!", title="Dynamic Toggle")
        #If button is true then set it to false
        elif dynamicToggle == True:
            dynamicToggle = False
            window['TOGGLE'].Update(text='Off', button_color='white on red')
            sg.popup_ok("Dynamic toggle is Off! Any organizations added/removed will NOT be automatically saved to excel!", title="Dynamic Toggle")
        print("The user has switched the dynamic toggle to " + str(dynamicToggle) + "!")
    #Event looks for table and clicked to show information on organization clicked
    elif event is not None and ('-TABLE-' and '+CLICKED+' in event) and (event[2][0] != None):
        getOrganizationPopup((getOrganizationFromClick(event, partners)), partnerInformation, collectedInformation, dynamicToggle)
    elif event is not None and ('VIEW' in event):
    #Opens the view information window which shows user's collected partners/information
        ViewInformationWindow(collectedInformation, dynamicToggle)
    #Changes table based on filter or shows FAQs from the menu bar
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