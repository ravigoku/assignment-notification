# assignment-notification

''Created on Jun 13, 2016
 
@author: Ravi'''

""" Comments
* need to set client name
* need to set client number
"""



#===============================================================================
# Imports
#===============================================================================
from twilio.rest import TwilioRestClient
import datetime
import openpyxl

#===============================================================================
# Functions
#===============================================================================
def sendReminderText(eventType,subject,dayOfWeek,clientCellPhone):
    """ sends a reminder text to client given: type, subject, day of week, and client cell number """
    if eventType == 'exam' or 'quiz':
        message = "Hey, your " +subject+" exam is coming up next "+dayOfWeek+". Text "+userName+", giving him the days you will study and a brief description of how you're going to study."
    elif eventType == 'project':
        message = "Hey, your " +subject+" project is due on the "+dayOfWeek+". Text "+userName+", giving him a brief description of the project. The days you will work on it, and specifically what you will do on those days." 
    else:
        message = message = "Hey, your " +subject+" Exam is coming up next "+dayOfWeek+". Text "+userName+", and let him know when you plan on finishing it."
    twilioCli.messages.create(body=message,from_=myTwilioNumber, to=clientCellPhone)
    print("client message sent.")
    
def sendNotificationText(eventType,name,subject,dayOfWeek,myCellPhone):
    """ Sends weekly notification text to user given: type, name, subject, day of week, user cell number """
    message = name+" has a(n) "+subject+""+eventType+" due "+dayOfWeek+"."
    twilioCli.messages.create(body=message,from_=myTwilioNumber, to=myCellPhone) 
    print("user message sent.")

def rowToDate (row):
    """ Insert a row containing ('Class','Month','Day','Year'),
and output a datetime object from 'Month','Day, and 'Year') """
    
    dateString = str(row[1].value) + ' ' + str(row[2].value) + ', ' + str(row[3].value)
    dateFromRow = datetime.datetime.strptime(dateString,'%B %d, %Y')
    print("date converted.\n")
    return dateFromRow

#===============================================================================
# Twilio Information
#===============================================================================

accountSID = 'AC6943252ee143e4ffb3c6539041e6073a'
authToken = 'ba86aa1a35c6e74b227349128e65fec0'

myTwilioNumber = '+17329976014'
twilioCli = TwilioRestClient(accountSID,authToken)

#===============================================================================
# Fixed Program Variables
#===============================================================================
clientName = "Pat"
clientCellPhone = "+17323280605"
userName= "Ravi"
myCellPhone = '+17323280605'

#===============================================================================
# Set Dates
#===============================================================================
now = datetime.datetime.now()
sevenDays = datetime.timedelta(days=7)
fourDays = datetime.timedelta(days=4)

eventDay7 = now + sevenDays
eventDay4 = now + fourDays

#===============================================================================
# Open Notification Workbook
#===============================================================================
wb = openpyxl.load_workbook('C:/Users/Ravi/Documents/RU 2016/Summer16/Academic Notification/notificationWorkbook.xlsx')

#===============================================================================
# Checking for Upcoming Exams 
#===============================================================================

examSheet = wb.get_sheet_by_name('Exams')
pair = tuple(examSheet['A2':'D9'])
for data in pair:
    examObj = rowToDate(data)
    print("Upcoming Exam...")
    
    if (examObj.month == now.month and examObj.day == now.day):
        print("true.")
        
        eventType = 'Exam'
        dayOfWeek = examObj.day
        subject = str(data[0].value)
        sendReminderText(eventType,subject,dayOfWeek,clientCellPhone)
        
    else:
        print("false.")
        
#===============================================================================
# Checking for Upcoming Quizzes
#===============================================================================
quizSheet = wb.get_sheet_by_name('Quiz')
pair = tuple(quizSheet['A2':'D9'])
for data in pair:
    quizObj = rowToDate(data)
    print("Upcoming Quiz...")
    
    if (quizObj.month == now.month and quizObj.day == now.day):
        print("true.")
        
        eventType = "Quiz"
        dayOfWeek = quizObj.day
        subject = str(data[0].value)
        
        sendReminderText(eventType,subject,dayOfWeek,clientCellPhone)
        
    else:
        print("false.")
        
#===============================================================================
# Checking for Upcoming Projects
#===============================================================================
projectSheet = wb.get_sheet_by_name('Projects')
pair = tuple(projectSheet['A2':'D9'])
for data in pair:
    projObj = rowToDate(data)
    print("Upcoming Project...")
    
    if (projObj.month == now.month and projObj.day == now.day):
        print("true.")
        
        eventType = 'Project'
        dayOfWeek = projObj.day
        subject = str(data[0].value)
        
        sendReminderText(eventType,subject,dayOfWeek,clientCellPhone)
        
    else:
        print("false.")
        
#===============================================================================
# Checking for Upcoming Assignments
#===============================================================================
weeklyAssSheet = wb.get_sheet_by_name('Weekly Assignments')
pair = tuple(weeklyAssSheet['A2':'D9'])
for data in pair:
    assObj = rowToDate(data)
    print("Upcoming Weekly Assignment...")
    
    if (assObj.month == now.month and assObj.day == now.day):
        print("true.")
        
        eventType = 'Assignment'
        dayOfWeek = assObj.day
        subject = str(data[0].value)
        
        sendReminderText(eventType,subject,dayOfWeek,clientCellPhone)
        
    else:
        print("false.")
