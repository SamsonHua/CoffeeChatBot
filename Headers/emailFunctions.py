#Import windows 32 client
import os
import win32com.client as client
import Headers.csvProcessing as csvProcessing
import numpy
import datetime


#==============================================#
#                ATTA FIKA BOT                 #
#==============================================#
#       DEVELOPED BY CONRAD B & SAMSON H       #
#==============================================#
#    YOU MUST LOGIN TO FIKA@ATTABOTICS.COM     #
#               FOR THIS TO WORK               #
#==============================================#

#Create Outlook Object
outlook = client.Dispatch('Outlook.Application')

#Set it so that it sends from fika@attabotics.com
#FikaBot sends a fancy HTML formatted email, so the email is first imported. "email.html is the 2 person match" while "email2.html" is the 3 person match since I didn't figure out dynamic html message sending in time
account = outlook.Session.Accounts['fika@attabotics.com']
with open('Headers/email.html', 'r') as myfile:
    data=myfile.read()

with open('Headers/email2.html', 'r') as myfile2:
    data2=myfile2.read()

#============================================================#
# >>> HELPER FUNCTION TO ACTUALLY SEND THE EMAIL
#============================================================#
def sendEmail(name1, name2, email1, email2):
    #Create outlook item 0 which is an email
    meeting = outlook.CreateItem(0)
    meeting.To = email1 + ';' + email2
    #Set parameters for meeting request
    meeting.Subject = "[Fika] - Your match for this week! \u2615 Friday Fika with " + name1 + " and " + name2 + "!"
    meeting.HTMLBody = data
    meeting.SendUsingAccount = account
    meeting.Display()
    # meeting.Save()
    # meeting.Send()


#============================================================#
# >>> HELPER FUNCTION TO SEND TRIPLE EMAIL
#============================================================#

def sendTripleEmail(name1, name2, name3, email1, email2, email3):
    #Create outlook item 0 which is an email
    meeting = outlook.CreateItem(0)
    meeting.To = email1 + ';' + email2 + ';' + email3
    #Set parameters for meeting request
    meeting.Subject = "[Fika] - Your match for this week! \u2615 Friday Fika with " + name1 + ", " + name2 + ", and " + name3 + "!"
    meeting.HTMLBody = data2
    meeting.SendUsingAccount = account
    meeting.Display()
    # meeting.Save()
    # meeting.Send()

#============================================================#
# >>> HELPER FUNCTION TO PARSE FIRST NAME
#============================================================#

def parseFirstName(email):
    occurences = email.count(".")

    if occurences > 1:
        firstName = email.split('.')[0]
        firstName = firstName.capitalize()
    else:
        firstName = email.split('@')[0]
        firstName = firstName.capitalize()
    return firstName

#============================================================#
# >>> FUNCTION TO SEND EMAILS TO EVERYONE IN "ThisWeeksMatches.xlsx"
#============================================================#

def sendAllEmails():
    #Import CSV file
    Pairings = csvProcessing.getPairings()

    for rows in Pairings:
        if rows[0] == "NULL" and rows[1] == "NULL" and rows[2] == "NULL":
            continue
        if rows[2] == "NULL":
            name1 = parseFirstName(rows[0])
            name2 = parseFirstName(rows[1])
            email1 = rows[0]
            email2 = rows[1]
            sendEmail(name1,name2,email1,email2)
            # print("Pairing " + name1 + " and " + name2 + " Emails: " + email1 + " and " + email2)
        else:
            name1 = parseFirstName(rows[0])
            name2 = parseFirstName(rows[1])
            name3 = parseFirstName(rows[2])
            email1 = rows[0]
            email2 = rows[1]
            email3 = rows[2]
            sendTripleEmail(name1,name2,name3,email1,email2,email3)
            # print("OOOOH BABY A TRIPLE, Pairing " + name1 + " and " + name2 + " and " + name3 + " Emails: " + email1 + " and " + email2 + " and " + email3)


