'''Created on Jun 13, 2016
 
@author: Ravi'''

''' assignment-notification v1.0 '''

#===============================================================================
# Imports
#===============================================================================
from twilio.rest import TwilioRestClient
import datetime, openpyxl, re

#===============================================================================
# Functions
#===============================================================================
def sendReminderText(eventType,subject,dayOfWeek,clientCellPhone):
    eventType = eventType.lower()
    """ sends a reminder text to client given: type, subject, day of week, and client cell number """
    if eventType == 'exam' or eventType =='quiz':
        message = "Hey, your " +subject+" exam is coming up next "+ str(dayOfWeek)+". Text "+userName+", giving him the days you will study and a brief description of how you're going to study."
    elif eventType == 'project':
        message = "Hey, your " +subject+" project is due on the "+str(dayOfWeek)+". Text "+userName+", giving him a brief description of the project. The days you will work on it, and specifically what you will do on those days." 
    else:
        message = "Hey, your " +subject+" weekly assignment is coming up next "+str(dayOfWeek)+". Text "+userName+", and let him know when you plan on finishing it."
    twilioCli.messages.create(body=message,from_=myTwilioNumber, to=clientCellPhone)
    print("client message sent.\n")

def rowToDate (row):
    """ Insert a row containing ('Class','Month','Day','Year'),
    and output a datetime object from 'Month','Day, and 'Year') """
    
    dateString = str(row[1].value) + ' ' + str(row[2].value) + ', ' + str(row[3].value)
    dateFromRow = datetime.datetime.strptime(dateString,'%B %d, %Y')
    print("date converted.")
    return dateFromRow

def checkSheet(wb,clientCellPhone,nameOfSheet,isAssignment):
    nameOfSheet = str(nameOfSheet)
    sheet = wb.get_sheet_by_name(nameOfSheet)
    pair = tuple(sheet['A2':'D9'])
    if isAssignment == False:
        dayToCheck = datetime.datetime.now() + datetime.timedelta(days=7)
        
    if isAssignment == True:
        dayToCheck = datetime.datetime.now() + datetime.timedelta(days=4)
    
    for data in pair:
        dateObj = rowToDate(data)
        print("Upcoming "+nameOfSheet+"...")
        
        if (dateObj.month == dayToCheck.month and dateObj.day == dayToCheck.day):
            print("true.")
            
            eventType = nameOfSheet
            dayOfWeek = dateObj.strftime('%a')
            subject = str(data[0].value).lower()
            
            sendReminderText(eventType,subject,dayOfWeek,clientCellPhone)
            
        else:
            print("false.")

def setCellNumber(wb):
    sheet = wb.get_sheet_by_name('Exam')
    return '+1'+str(sheet['F2'].value)

def setName(wb):
    sheet = wb.get_sheet_by_name('Exam')
    return str(sheet['E2'].value)

def activateTwilio(location):
    location = str(location)
    f = open(location)
    key = f.read()
    SIDRegex = re.compile(r'(accountSID = )(.*)')
    accountSID = SIDRegex.search(key).group(2)
    
    authRegex = re.compile(r'(authToken = )(.*)')
    authToken = authRegex.search(key).group(2)
    
    TwilioNumRegex = re.compile(r'(myTwilioNumber = )(.*)')
    myTwilioNumber = TwilioNumRegex.search(key).group(2)
    
    myCellRegex = re.compile(r'(myCellPhone = )(.*)')
    myCellPhone = myCellRegex.search(key).group(2)
    
    return accountSID,authToken,myTwilioNumber,myCellPhone

def startProcess(fileArray):
    for f in fileArray:
        wb = openpyxl.load_workbook(f)
        clientCellPhone = setCellNumber(wb)
        checkSheet(wb,clientCellPhone,'Exam',False)
        checkSheet(wb,clientCellPhone,'Quiz',False)
        checkSheet(wb,clientCellPhone,'Project',False)      
        checkSheet(wb,clientCellPhone,'Weekly Assignment', True)

def startWeeklyProcess(fileArray):
    message = ''
    for f in fileArray:
        wb = openpyxl.load_workbook(f)
        name = name = setName(wb)
        message = checkWeeklySheet(wb,name,message,'Exam')
        message = checkWeeklySheet(wb,name,message,'Quiz')
        message = checkWeeklySheet(wb,name,message,'Project')      
        message = checkWeeklySheet(wb,name,message,'Weekly Assignment')
        
    if message != '':
        twilioCli.messages.create(body=message,from_=myTwilioNumber, to=myCellPhone)
        print("user message sent.\n")              

def checkWeeklySheet(wb,name,message,nameOfSheet):
    nameOfSheet = str(nameOfSheet)
    sheet = wb.get_sheet_by_name(nameOfSheet)
    
    pair = tuple(sheet['A2':'D9'])

    dayToCheck = datetime.datetime.now() + datetime.timedelta(days=8)

    for data in pair:
        dateObj = rowToDate(data)
        print("Upcoming "+nameOfSheet+"...")
        
        if (dateObj>=datetime.datetime.now() and dateObj<=dayToCheck):
            print("true.")
        
            eventType = nameOfSheet
            dayOfWeek = dateObj.strftime('%a')
            subject = str(data[0].value).lower()
            message = message +"*"+name+' has a '+subject+' '+eventType.lower()+' due '+dayOfWeek+'.\n'
            print('message updated.\n')
        else:
            print("false.\n")
        
    return message


#===============================================================================
# main
#===============================================================================

# twilio info

accountSID,authToken,myTwilioNumber,myCellPhone = activateTwilio('C:\\Users\\Ravi\\Documents\\RU 2016\\Summer16\\Academic Notification\\twilioInformation.txt')
twilioCli = TwilioRestClient(accountSID,authToken)

userName= "Ravi"

fileArray = ['C:/Users/Ravi/Documents/RU 2016/Summer16/Academic Notification/notificationWorkbook.xlsx']

# executed everyday
startProcess(fileArray)

# executed Sundays
if (datetime.datetime.now().strftime('%w') == '0'):
    startWeeklyProcess(fileArray)
