#Project 3 - Group 3
#Date: 2023-03-09
#Activities to run 
# 1 - Excel automation
# 2 - Get PC and User Info
# 3 - Encrypting and decrypting data

import os, openpyxl, time
import win32com.client as win32
import  win32api, win32process, win32com, win32crypt, win32file

def menuSelection():
    userInfoPrint = """
    1 = Excel Automation
    2 = Get PC and System Info
    3 = Encrypting and decrypting data
    0 = Exit
    SELECT MENU: """
    #Select loop menu 
    while True:
        try:
            userInput = int(input(userInfoPrint))
            if userInput not in range(0,5):
                print("Input not Available>>>Try Again")
                
            if userInput == 0:
                print("Exiting...Thank you for using this program")
                break
            
            elif userInput == 1:
                excelAutomation(excelFile)
                time.sleep(5)
            
            elif userInput == 2:
                getUserandSystemInfo()
                time.sleep(5)
                
            elif userInput == 3:
                EncryptingAndDecrypting()
                time.sleep(5)
        
        except ValueError:
            print("Invalid Input>>> Input an Interger available")
            print("WAIT FOR 10S")
            time.sleep(10)
            
        
######################################################################################        
#####################################################################################
#EXCEL AUTOMATION
"""
This excel automation is basically working with database (generated from mockaroo)
You could add, remove or change the data in the database
You could calculate in the database

Process
1 - Import documentation from pywin32
2 - Open an excel worksheet
3 - select a worksheet and column
4 - create a new row and calculate the sum of each employee earned within their work years
5 - save excel file and close the excel file
"""
excelFile = "C:/Users/josep/Documents/Assignment 3/mockData.xlsx"

#Function for activity One
def excelAutomation(excelFile):
    
    #open excel file
    openExcel = win32com.client.Dispatch("Excel.Application")

    #Open the excel file called "mockData" and worksheet "data"
  
    workbook = openExcel.Workbooks.Open(excelFile)
    
    #selecting the worksheet from excel
    worksheet = workbook.Worksheets("data")
    
    """
        Calculating the annual_income * years of experience to get total earned in working years
        i.e annual_income * years_of_experience = total earned in working years
        """

    #create a new row and add a header to it. In the excel document I is our new row
    worksheet.Range("I1").Value = "Total_Earned_in_Working_Years"
    
    #Add formula
    worksheet.Range("I2").Formula = "=G2*H2"
    lastRow = worksheet.Cells(worksheet.Rows.Count, 1).End(win32.constants.xlUp).Row
    for row in range(2, lastRow + 1):
        worksheet.Range("I" + str(row)).Formula = "=G" + str(row) + "*H" + str(row)
    
    #Input Menu for excel
    userInputInfo = """
    EXCEL AUTOMATION
    1 - Adding Data to the database
    2 - Removing Data from the database
    3 - Searching Data from the database
    0 - Exit
    SELECT MENU: """
    
    #Error handling on excel automation input
    while True:
        try:
            userInput = int(input(userInputInfo))
            if userInput not in range(0, 4):
                print("Number not in selection")

            if userInput == 0:
                print("Exiting Program")
                break
                
            if userInput == 1:
                # adding data into excel file
                while True:
                    try:
                        """
                        when inputting the account number a zero is not counted
                        for instance my account number is 000788570, the first three zero is not counted
                        so it is actually 788570. Maybe the excel program is the cause of this error or mackaroo
                        we generated online
                        """
                        account_number = int(input("Enter account number: "))
                        if len(str(account_number)) != 8:
                            print("Invalid account number: must be an integer with 8 digits")
                            pass
                        else:
                            break
                    except ValueError:
                        print("Invalid input: account number must be an integer")

                first_name = str(input("Enter first name: "))
                last_name = str(input("Enter last name: "))
                email = str(input("Enter email: "))
                city = str(input("Enter city: "))
                job = str(input("Enter job: "))
                annual_income = float(input("Enter annual income in $: "))
                years_of_experience = int(input("Enter years of experience: "))

                # find the last row of the data
                lastRow = worksheet.Cells(worksheet.Rows.Count, 1).End(-4162).Row

                # write data to the next row
                worksheet.Cells(lastRow + 1, 1).Value = account_number
                worksheet.Cells(lastRow + 1, 2).Value = first_name
                worksheet.Cells(lastRow + 1, 3).Value = last_name
                worksheet.Cells(lastRow + 1, 4).Value = email
                worksheet.Cells(lastRow + 1, 5).Value = city
                worksheet.Cells(lastRow + 1, 6).Value = job
                worksheet.Cells(lastRow + 1, 7).Value = annual_income
                worksheet.Cells(lastRow + 1, 8).Value = years_of_experience
            
                print("DATA ADDED TO EXCEL FILE")
            #For removing data using account number 
            elif userInput == 2:
                RemoveAccount = int(input("Input account number to remove: ")) 
                removeRow = None
                for row in range(1, worksheet.UsedRange.Rows.Count + 1):
                    if worksheet.Cells(row, 1).Value == RemoveAccount:
                        removeRow = row
                        time.sleep(5)
                        print("DATA DON'T EXIST")
                        break

                if removeRow is not None:
                    # Erase data in the row
                    worksheet.Range(f"A{removeRow}:H{removeRow}").Value = None
                    print("DATA REMOVED FROM EXCEL")
                    
                    # Automatically shift data from empty space to the end of the row
                    worksheet.Range(f"A{removeRow}:H{worksheet.UsedRange.Rows.Count}").Cut(Destination=worksheet.Range
                                                                                                (f"A{removeRow}:H{removeRow-1 + worksheet.UsedRange.Rows.Count-removeRow+1}"))
            #SEARCHING Account
            elif userInput == 3:
                #employee should search based on the account number since they might have same first and last name
                #equal but diffrent account number
                account_number = int(input("Enter account number: "))
                email = str(input("Enter email: "))
                query_accountNumber = worksheet.Columns(1).Find(account_number)          
                query_email = worksheet.Columns(4).Find(email)  
                if query_accountNumber is not None and query_email is not None:
                    first_name = query_accountNumber.Offset(1, 2).Value
                    last_name = query_accountNumber.Offset(1, 3).Value
                    print(f"Available from the name: {first_name} {last_name}")
                else:    
                    print("Data not found")     
        except ValueError:
            print("Invalid Input>>> Input an Interger available")
            
    
    #sorting the rows in alphabetical order with their data
    sortingRange = worksheet.Range('A1:I1' + str(lastRow))
    lastRow = worksheet.Cells(worksheet.Rows.Count, "B").End(-4162).Row
    sortingRange.Sort(Key1=sortingRange.Columns(2), Order1=1, Orientation=1, Header=1)


    workbook.Save()
    workbook.Close()
######################################################################################

#Encrypting and decrypting data
######################################################################################
#The process of encrypting data basically is to protect the data for cyberattackers as well as
#safeguard the data during file transfer. However it's always a good practice to protect the data
"""
1 - Import the unencrypted file from the directory
2 - Read the unencrypted data from the file
3 - Encrypt the data using the method win32crypt.CryptProtectData
4 - Write the encrypted data to the file
5 - File Transfer and ready for decryption
!!! The easiest file type to encrypt and decrypt are txt. files, pdfs and microsoft files (docx, xlsx)
can be encrypted and decrypted. However most of times they easily get corrupt thus a loss of data!!!

"""
#Function for encrypting and decrypting data
def EncryptingAndDecrypting():
    selectProcessInfo = """
    ENCRYPTING AND DECRYPTING DATA
    INPUT YOUR CHOICE
    1 - ENCRYPTION
    2 - DECRYPTION
    0 - EXIT
    SELECT MENU: """
    
    selectProcess = int(input(selectProcessInfo))
    if selectProcess == 1:
        fileNameInput = input("Enter the file name: ")
        fileDirectory = f"C:/Users/josep/Documents/Assignment 3/encryptionFolder/{fileNameInput}"
        encryptedDataFile = f"C:/Users/josep/Documents/Assignment 3/decryptionFolder/enc_{fileNameInput}"
        
        if os.path.exists(fileDirectory):
            print("The file exists and ready for encryption>>>>")
            
            with open(fileDirectory, 'rb') as file:
                readData = file.read()
            
                #Try and error handling Encryption
            try:
                encryptingData = win32crypt.CryptProtectData(readData, DataDescr=None,
                                                            OptionalEntropy=None,
                                                            Reserved=None,
                                                            PromptStruct=None,
                                                            Flags=0)
            except:
                print("Error in data encryption")

            #Writing encrypting data to new one
            with open(encryptedDataFile, "wb") as file:
                file.write(encryptingData)
                print("Data encrypted")
                
                #Moving file to Decryption folder 
                
                print("File moved to Decryption folder")
              
        if not os.path.exists(fileDirectory):
            print("The file does not exist")
                
    elif selectProcess == 2: 
        userInput = input("Enter the name of the file you want to decrypt: ")
        encryptedFolder = f"C:/Users/josep/Documents/Assignment 3/decryptionFolder/{userInput}"
        if os.path.exists(encryptedFolder):  
            print("The file exists and is ready for decryption.")
            time.sleep(5)
            with open(encryptedFolder, "rb") as file:
                readData = file.read()
            try:
                decryptingData = win32crypt.CryptUnprotectData(readData, OptionalEntropy=None,
                                                                Reserved=None, 
                                                                PromptStruct=None, 
                                                                Flags=0)[1]
            except:
                print("Error in data decryption")
            newDecryptedFile = f"C:/Users/josep/Documents/Assignment 3/decryptionFolder/dec_{userInput}"
            
            with open(newDecryptedFile, "wb") as file:
                file.write(decryptingData)
                print("Data has been decrypted")
                
        if not os.path.exists(encryptedFolder):
            print("The file does not exist")
            
    elif selectProcess == 0:
        print("Exiting program")
              
######################################################################################

#GET SYSTEM INFO
######################################################################################
"""
We get all PC user info and its system which includes username, logical devices, Microsoft version,
etc... and print to console. Normally we can get the username and the system information for IT technician
to get system Information since it is harder to get some of the system information from the windows system
"""
def getUserandSystemInfo():
    #Get User Info
    userInfoInput = """
    System Info
    1 - Get Windows Current version
    2 - Get System Cache Info
    3 - Get Current Process of the memory 
    4 - Get Info on Current Threads and Proccesors
    5 - Print full data into file except Memory Info
    6 - Get Printed Memory Info
    0 - Exit
    Select Menu: """
    #USERNAME
    getUserInfo = win32api.GetUserName()
    
    #DEFINING CURRENT WINDOWS VERSION
    """
    1 - Major Version
    2 - Minor Version
    3 - Build Version
    4 - Platform ID"""
    currentWindowVersion = win32api.GetVersionEx()
    getMajorVersion = currentWindowVersion[0]
    getMinorVersion = currentWindowVersion[1]
    getBuildVersion = currentWindowVersion[2]
    getPlatformId = currentWindowVersion[3]
    
    #Getting cache file size
    getsystemCacheFileSize = win32api.GetSystemFileCacheSize()
    maxCache = getsystemCacheFileSize[0]
    minCache = getsystemCacheFileSize[1]
    Flags = getsystemCacheFileSize[2]
    
    # print(getMemoryInfo)
    # getProcess = win32api.GetCurrentProcess()
    # getMemoryInfo = win32process.GetProcessMemoryInfo(getProcess)
    # print(getMemoryInfo)
    getProcess = win32api.GetCurrentProcess()
    getMemoryInfo = win32process.GetProcessMemoryInfo(getProcess)
    # print(getMemoryInfo)
    for i, j in getMemoryInfo.items():
        printInfo = (f"{i}: {j}\n")
        
     #Getting current thread
    getcurrentThread =win32api.GetCurrentThread()
    getCurrentThreadId = win32api.GetCurrentThreadId()
    
    #Getting Processors Name
    wmi = win32com.client.GetObject("winmgmts:")
    processors = wmi.ExecQuery("Select * from Win32_Processor")
    for system in processors:
        NameProccesorsandCloudSpeed = f"Proccesor Name and Clock Speed: {system.Name.strip()} {system.MaxClockSpeed}MHz"
        NumberofCores = f"Number of Cores: {system.NumberOfCores}"
        """print( f"Proccesor Name and Clock Speed: {system.Name.strip()} {system.MaxClockSpeed}MHz")
        print(f"Number of system core: {system.NumberOfCores} Cores")"""
        
    ProcessorsInfo = win32api.GetSystemInfo()
    ProcessorArchitecture = ProcessorsInfo[0]
    pageSize = ProcessorsInfo[1]
    proccesorsType = ProcessorsInfo[7]
    processorsLevel = ProcessorsInfo[1]
    numberOfProccesors = ProcessorsInfo[5]

    while True:
        try:
            userInput = int(input(userInfoInput))
            if userInput not in range(0, 7):
                print("Input not available_Try Again")
                
            elif userInput == 0:
                break
            
            elif userInput == 1:
                # currentWindowVersion = win32api.GetVersionEx()
                # getMajorVersion = currentWindowVersion[0]
                # getMinorVersion = currentWindowVersion[1]
                # getBuildVersion = currentWindowVersion[2]
                # getPlatformId = currentWindowVersion[3]
                
                print(f"Major Version: {getMajorVersion}")
                print(f"Minor Version: {getMinorVersion}")
                print(f"Build Version: {getBuildVersion}")
                print(f"Platform Id: {getPlatformId}")
                
            elif userInput == 2:
                #Cache size info
                # getsystemCacheFileSize = win32api.GetSystemFileCacheSize()
                # maxCache = getsystemCacheFileSize[0]
                # minCache = getsystemCacheFileSize[1]
                # Flags = getsystemCacheFileSize[2]
                
                print(f"User: {getUserInfo}")       
                print(f"MaximumCache: {maxCache}")
                print(f"MinimumCache: {minCache}")
                print(f"Flags: {Flags}")
            
            elif userInput == 3:
                # getProcess = win32api.GetCurrentProcess()
                # getMemoryInfo = win32process.GetProcessMemoryInfo(getProcess)
                # # print(getMemoryInfo)
                # for i, j in getMemoryInfo.items():
                #     printInfo = (f"{i}: {j}\n")
                    print(printInfo)
            
            elif userInput == 4:
                """getCurrentThreadId = win32api.GetCurrentThreadId()
                print(getCurrentThreadId)
                """
               
                #Selective printing the current thread
                print(f"Current Thread Id: {getCurrentThreadId}")
                print(f"Current Thread: {getcurrentThread}")
                
                #Getting Proccesors system Info
                ProcessorsInfo = win32api.GetSystemInfo()
                """
                From the documentation, This system info include
                :
                1 - ProcessorArchitecture
                2 - Page Size
                3 - lp Minimum Application Address
                4 - lp Maximum Application Address
                5 - dw Active Processor Mask
                6 - Number Of Processors
                7 - Processor Type
                8 - Allocation Granularity
                9 - Processor Level ,Processor Revision)
                                    """
                ProcessorArchitecture = ProcessorsInfo[0]
                pageSize = ProcessorsInfo[1]
                proccesorsType = ProcessorsInfo[7]
                processorsLevel = ProcessorsInfo[1]
                numberOfProccesors = ProcessorsInfo[5]
                
                                    
                #Getting Processors Name
                wmi = win32com.client.GetObject("winmgmts:")
                processors = wmi.ExecQuery("Select * from Win32_Processor")
                for system in processors:
                    NameProccesorsandCloudSpeed = f"Proccesor Name and Clock Speed: {system.Name.strip()} {system.MaxClockSpeed}MHz"
                    NumberofCores = f"Number of Cores: {system.NumberOfCores}"
                    """print( f"Proccesor Name and Clock Speed: {system.Name.strip()} {system.MaxClockSpeed}MHz")
                    print(f"Number of system core: {system.NumberOfCores} Cores")"""
                    print(NameProccesorsandCloudSpeed)
                    print(NumberofCores)
                    # print(f" {system.NumberOfLogicalProcessors} Logical Processors")
                    # print(f" {system.NumberOfPhysicalProcessors} Physical Processors")
                    # print(f" {system.ProcessorType} Processor Type")
                    # print(f" {system.AllocationGranularity} Allocation Granularity")
                    # print(f" {system.ProcessorLevel} Processor Level")
                    # print(f" {system.ProcessorRevision} Processor Revision")
                print(f"Processor Architecture: {ProcessorArchitecture}")
                print(f"Page Size: {pageSize}")
                print(f"Proccesors Type: {proccesorsType}")
                print(f"Processors level: {processorsLevel}")
                print(f"Number Of Processors: {numberOfProccesors}")
            
            elif userInput == 5:
                #Printing system info into file apart from memory info    
                print('*'*40)
                infoBoundaries = "*" * 20
                writeData= f"""
                User Name: {getUserInfo}
                {infoBoundaries}
                WINDOWS CURRENT VERSION
                Major Version: {getMajorVersion}
                Minor Version: {getMinorVersion}
                Build Version: {getBuildVersion}
                Platform Id: {getPlatformId}
                {infoBoundaries}
                SYSTEM CACHE INFO
                MaximumCache: {maxCache}
                MinimumCache: {minCache}
                Flags: {Flags}
                {infoBoundaries}
                INFO ON THREADS AND PROCCESORS
                {NameProccesorsandCloudSpeed}
                {NumberofCores}
                Processor Architecture: {ProcessorArchitecture}
                Page Size: {pageSize}
                Proccesors Type: {proccesorsType}
                Processors level: {processorsLevel}
                Number Of Processors: {numberOfProccesors}
                {infoBoundaries}
                 """
                
                try:
                    fileHandle = win32file.CreateFile(getUserInfo + "systemInfo.txt",
                                                        win32file.GENERIC_WRITE,
                                       win32file.FILE_SHARE_WRITE,
                                  None, win32file.CREATE_ALWAYS, 
                                  0, None)
                    
                    writeData = writeData.encode('ascii')
                    win32file.WriteFile(fileHandle, writeData)
                    print("Data written Succesffully")
                except:
                    print("Error writing system info to file")
                
                win32file.CloseHandle(fileHandle)
            
            elif userInput == 6:
                #Printing memory info
                print('*'*40)
                try:
                    fileHandle = win32file.CreateFile(getUserInfo + "memoryInfo.txt",
                                                      win32file.GENERIC_WRITE,
                                       win32file.FILE_SHARE_WRITE,
                                  None, win32file.CREATE_ALWAYS, 
                                  0, None)
                    for i, j in getMemoryInfo.items():
                        printInfo = (f"{i}: {j}\n")
                        writeData = printInfo.encode('ascii')
                        win32file.WriteFile(fileHandle, writeData)
                        print("Data written Succesffully")
                    win32file.CloseHandle(fileHandle)
                except:
                    print("Error writing system info to file")

        except ValueError:
            print("Invalid Input>>> Input an Interger available")
                
######################################################################################
#CALL MENU SELECT FUNCTION
menuSelection()