'''Created on Jun 13, 2016
 
@author: Ravi'''

""" Comments
* need to set client name
* need to set client number
"""

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
    
def sendNotificationText(eventType,name,subject,dayOfWeek,myCellPhone):
    """ Sends weekly notification text to user given: type, name, subject, day of week, user cell number """
    message = name+" has a(n) "+subject+""+eventType+" due "+str(dayOfWeek)+"."
    twilioCli.messages.create(body=message,from_=myTwilioNumber, to=myCellPhone) 
    print("user message sent.\n")

def rowToDate (row):
    """ Insert a row containing ('Class','Month','Day','Year'),
    and output a datetime object from 'Month','Day, and 'Year') """
    
    dateString = str(row[1].value) + ' ' + str(row[2].value) + ', ' + str(row[3].value)
    dateFromRow = datetime.datetime.strptime(dateString,'%B %d, %Y')
    print("date converted.\n")
    return dateFromRow

def checkSheet(nameOfSheet,Assignment):
    nameOfSheet = str(nameOfSheet)
    sheet = wb.get_sheet_by_name(nameOfSheet)
    pair = tuple(sheet['A2':'D9'])
    if Assignment == False:
        dayToCheck = datetime.datetime.now() + datetime.timedelta(days=7)
        
    if Assignment == True:
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

#===============================================================================
# Twilio Information
#===============================================================================


accountSID,authToken,myTwilioNumber,myCellPhone = activateTwilio('C:\\Users\\Ravi\\Documents\\RU 2016\\Summer16\\twilioInformation.txt')
twilioCli = TwilioRestClient(accountSID,authToken)

#===============================================================================
# Fixed Program Variables
#===============================================================================
clientName = "Pat"
clientCellPhone = "+17323280605"
userName= "Ravi"


#===============================================================================
# Open Notification Workbook
#===============================================================================
wb = openpyxl.load_workbook('C:/Users/Ravi/Documents/RU 2016/Summer16/Academic Notification/notificationWorkbook.xlsx')

checkSheet('Exam',False)
checkSheet('Quiz',False)
checkSheet('Project',False)      
checkSheet('Weekly Assignment', True)

