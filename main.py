import speech_recognition as sr
import pyttsx3
import imaplib
import email
import re
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


engine = pyttsx3.init()
recognizer = sr.Recognizer()

receive_mail_conn = None
send_mail_conn = None
emailId = ''
OFFSET=5

def saySomething(phrase):
    print('Engine says...', phrase)
    engine.say(phrase)
    engine.runAndWait()


def getUserInput(stmt, listen_for=None, confirm_input=True, doNotRetry=False, remove_spaces=False, to_lower_case=False):
    with sr.Microphone() as source:
        recognizer.adjust_for_ambient_noise(source, duration=1)
        if (stmt):
            saySomething(stmt)
        print('Listening...')
        recordedaudio = recognizer.listen(source, None, listen_for)
    try:
        userInput = recognizer.recognize_google(recordedaudio, language='en-US')
        # print('userInput 11', userInput)
        if (remove_spaces):
            userInput = ''.join(userInput.split(' '))
        if (to_lower_case):
            userInput = userInput.lower()
        print('User said...', userInput)
        if (confirm_input):
            saySomething('You said, ' + userInput)
        return userInput
    except Exception as ex:
        print('Error while getting user input...', ex)
        if doNotRetry:
            handleFailure()
        else:
            return getUserInput(stmt, doNotRetry=True)


def sendMail(receipient, mail):
    global send_mail_conn
    global emailId
    print('Sending email...')
    try:
        send_mail_conn.sendmail(emailId, receipient, mail.as_string())
    except Exception as ex:
        print('Error while getting sending mail...', ex)
        handleFailure()
        return


def handleFailure():
    saySomething('Unexpected error occured! Please try again later!')
    exit()


def notYetConfigured():
    saySomething('This option is not yet configured! Please try again later!')


def composeMail():
    global emailId
    # COMPOSE EMAIL
    # -----------------
    # Ask for recipient
    receipient = getUserInput('Please provide the receipient\'s email address.', remove_spaces=True, to_lower_case=True)
    # Ask for subject
    subject = getUserInput('Please provide the subject of the email.')
    # Ask for body
    body = getUserInput('Please provide the content of the email.', 10)
    # Get confirmation
    confirmation = getUserInput('Are you sure you want to send the email?')
    if (confirmation == 'yes'):
        # Send Email
        saySomething('Please wait while the mail is being sent.')
        mail = MIMEMultipart('alternative')
        mail['Subject'] = subject
        mail['From'] = emailId
        mail['To'] = receipient
        part1 = MIMEText(body, 'html')
        mail.attach(part1)
        sendMail(receipient, mail)
        saySomething('You mail has been sent successfully.')
    else:
        saySomething('You mail has been cancelled.')


def quitApp():
    saySomething('Closing the Application. Thanks for using.')
    quit()


def logout():
    saySomething('You have been logged-out')


def postLoginMenu():
    options = [
        'Compose mail',
        'Inbox',
        'Logout',
        'Quit'
    ]
    while True:
        saySomething('Here are your available options')
        for option in options:
            saySomething(option)
        choice = getUserInput('Please choose an option')
        if (choice == 'compose mail'):
            composeMail()
        elif (choice == 'inbox'):
            inbox()
        elif (choice == 'log out'):
            logout()
            return
        elif (choice == 'quit'):
            quitApp()
        else:
            saySomething('Could not recognize the option. Please try again')


def getUserInputLetterByLetter(stmt):
    saySomething(stmt)
    letters = []
    while True:
        user_input = getUserInput(None, confirm_input=False)
        if (user_input == 'done'):
            break
        else:
            letters.append(user_input)
    user_input = ''.join(letters)
    saySomething('You said, ' + ' '.join(letters))
    return user_input


def login():
    global emailId
    global send_mail_conn
    global receive_mail_conn
    # emailId = getUserInput('Please provide your email', remove_spaces=True, to_lower_case=True)
    # password = getUserInput('Please provide your password', remove_spaces=True, to_lower_case=True)
    emailId = 'naveenasai2001@gmail.com'
    password = 'tjictggamdbqfjli'
    # password = getUserInputLetterByLetter('Please provide your password')
    saySomething('Please wait while we login')
    try:
        SMTP_SERVER = "imap.gmail.com"
        SMTP_PORT = 993
        receive_mail_conn = imaplib.IMAP4_SSL(SMTP_SERVER)
        receive_mail_conn.login(emailId, password)
        send_mail_conn = smtplib.SMTP('smtp.gmail.com', 587) 
        send_mail_conn.ehlo()
        send_mail_conn.starttls()
        send_mail_conn.login(emailId,password)  
        saySomething('Login successful!')
        postLoginMenu()
    except Exception as ex:
        saySomething('Unable to login')


def parseEmail(data):
    for response_part in data:
        if isinstance(response_part, tuple):
            return email.message_from_bytes(response_part[1])


def getMailDetailsByEmailId(emailId):
    global receive_mail_conn
    receive_mail_conn.select('INBOX')
    category = "Primary"
    status, response = receive_mail_conn.uid('search', 'X-GM-RAW "category:' + category + '"')
    response = response[0].decode('utf-8').split()
    response.reverse()
    mail_checked = 0
    for uid in response:
        mail_checked = mail_checked+1
        status, data = receive_mail_conn.uid('fetch', uid, '(RFC822)')
        mail = parseEmail(data)
        print("Checking mail : " + mail['from'])
        if emailId.lower() in mail['from'].lower():
            return mail
        if mail_checked % 5 == 0:
            helper_text = "first"
            if mail_checked > 5:
                helper_text = "next"
            choice = getUserInput("couldn't find the mail from " + helper_text + " 5 mails in the inbox, should I continue the search?")
            if choice == "no":
                saySomething("stopping the search")
                return "canceled"
            else:
                saySomething("Continuing the search")
    return None
    

def getMails(skip=0, limit=OFFSET):
    global receive_mail_conn
    receive_mail_conn.select('INBOX')
    category = "Primary"
    status, response = receive_mail_conn.uid('search', 'X-GM-RAW "category:' + category + '"')
    response = response[0].decode('utf-8').split()
    response.reverse()
    response = response[skip:min(skip+limit, len(response))]
    mails = []
    for uid in response:
        status, data = receive_mail_conn.uid('fetch', uid, '(RFC822)')
        mail = parseEmail(data)
        mails.append(mail)
    return mails


def readMails(mails):
    saySomething("received mails from the following")
    for mail in mails:
        saySomething(mail["from"])


def extractContentFromHTML(html):
    content = re.sub('<[^>]*>', '', html)
    content = re.sub('\s+', ' ', content)
    return content


def readMailDetails(mail):
    saySomething('Here are the details of the mail from ' + mail['from'])
    content_type = mail.get_content_type()
    payload = mail.get_payload()
    if (content_type == 'text/html'):
        content = extractContentFromHTML(payload)
        saySomething('Warning! This content was extracted from HTML so it could be in-accurate.')
        saySomething(content)
    elif (content_type == 'multipart/alternative'):
        content = payload[0].get_payload()
        saySomething(content)
    else:
        print('ERROR! Unhandled content_type', content_type)
        saySomething('Could not fetch the email')


def inbox():
    options = [
        'read mail',
        'search mail',
        'go back',
    ]
    while True:
        saySomething('Here are your available options')
        for option in options:
            saySomething(option)
        choice = getUserInput('Please choose an option')
        if (choice == 'read mail'):
            readMailFromInbox()
        elif (choice == 'search mail'):
            from_address = getUserInput('Please provide a from address')
            mail_to_be_opened = getMailDetailsByEmailId(from_address)
            if mail_to_be_opened != "canceled":
                if mail_to_be_opened:
                    readMailDetails(mail_to_be_opened)
                else:
                    saySomething('Could not find email')
        elif (choice == 'go back'):
            break
        elif (choice == 'quit'):
            quitApp()
        else:
            saySomething('Could not recognize the option. Please try again')


def readMailFromInbox():
    saySomething("please wait while we load your mails from the inbox")
    skip = 0
    mails = getMails(skip)
    saySomething("Here are the latest " + str(OFFSET) + " mails from your inbox")
    readMails(mails)
    options = [
        'read more',
        'open mail',
        'go back'
    ]
    while True:
        saySomething('Here are your available options')
        for option in options:
            saySomething(option)
        choice = getUserInput('Please choose an option')
        if choice == "read more":
            saySomething("please wait")
            skip = skip + OFFSET
            mails = getMails(skip)
            readMails(mails)
        elif choice == "open mail":
            saySomething("please go back and navigate to search mail to open the mail")
        elif (choice == 'go back'):
            break
        elif (choice == 'quit'):
            quitApp()
        else:
            saySomething('Could not recognize the option. Please try again')


def mainMenu():
    saySomething('Welcome to Audio Email Service')
    options = [
        'Login',
        'Register',
        'Quit'
    ]
    while True:
        saySomething('Here are your available options')
        for option in options:
            saySomething(option)
        choice = getUserInput('Please choose an option')
        if (choice == 'login'):
            login()
        elif (choice == 'register'):
            notYetConfigured()
        elif (choice == 'quit'):
            quitApp()
        else:
            saySomething('Could not recognize the option. Please try again')


if __name__ == '__main__':
    mainMenu()
