import speech_recognition as sr
import yagmail
import pyttsx3
import imaplib
import email
import re
from bs4 import BeautifulSoup


engine = pyttsx3.init()
recognizer = sr.Recognizer()

mail_conn = None
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


def sendMail(receipient, subject, body):
    print('Sending email...')
    try:
        with open('.\configs\mail.xml') as f:
            content = f.read()
        config = BeautifulSoup(content, features='lxml-xml')
        email = str(config.mailConfig.email.contents[0])
        password = str(config.mailConfig.password.contents[0])
        sender = yagmail.SMTP(user=email, password=password)
        sender.send(to=receipient, subject=subject, contents=body)
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
    # COMPOSE EMAIL
    # -----------------
    # Ask for recipient
    receipient = getUserInput('Please provide the receipient\'s email address.')
    # Ask for subject
    subject = getUserInput('Please provide the subject of the email.')
    # Ask for body
    body = getUserInput('Please provide the content of the email.', 10)
    # Get confirmation
    confirmation = getUserInput('Are you sure you want to send the email?')
    if (confirmation == 'yes'):
        # Send Email
        saySomething('Please wait while the mail is being sent.')
        sendMail(receipient, subject, body)
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
    global mail_conn
    # emailId = getUserInput('Please provide your email', remove_spaces=True, to_lower_case=True)
    # password = getUserInput('Please provide your password', remove_spaces=True, to_lower_case=True)
    emailId = 'naveenasai2001@gmail.com'
    password = 'tjictggamdbqfjli'
    # password = getUserInputLetterByLetter('Please provide your password')
    saySomething('Please wait while we login')
    try:
        SMTP_SERVER = "imap.gmail.com"
        SMTP_PORT = 993
        mail_conn = imaplib.IMAP4_SSL(SMTP_SERVER)
        mail_conn.login(emailId, password)
        saySomething('Login successful!')
        postLoginMenu()
    except Exception as ex:
        saySomething('Unable to login')


def parseEmail(data):
    for response_part in data:
        if isinstance(response_part, tuple):
            return email.message_from_bytes(response_part[1])


def getMailDetailsByEmailId(emailId):
    global mail_conn
    mail_conn.select('INBOX')
    category = "Primary"
    status, response = mail_conn.uid('search', 'X-GM-RAW "category:' + category + '"')
    response = response[0].decode('utf-8').split()
    response.reverse()
    for uid in response:
        status, data = mail_conn.uid('fetch', uid, '(RFC822)')
        mail = parseEmail(data)
        if emailId.lower() in mail['from'].lower():
            return mail
    return None
    

def getMails(skip=0, limit=OFFSET):
    global mail_conn
    mail_conn.select('INBOX')
    category = "Primary"
    status, response = mail_conn.uid('search', 'X-GM-RAW "category:' + category + '"')
    response = response[0].decode('utf-8').split()
    response.reverse()
    response = response[skip:min(skip+limit, len(response))]
    mails = []
    for uid in response:
        status, data = mail_conn.uid('fetch', uid, '(RFC822)')
        mail = parseEmail(data)
        mails.append(mail)
    return mails


def readMails(mails):
    for mail in mails:
        saySomething("received mail from " + mail["from"])


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


def postInboxMenu():
    options = [
        'load more',
        'open mail',
        'go back',
    ]
    skip = OFFSET
    while True:
        saySomething('Here are your available options')
        for option in options:
            saySomething(option)
        choice = getUserInput('Please choose an option')
        if (choice == 'load more'):
            next_mails = getMails(skip)
            readMails(next_mails)
            skip += OFFSET
        elif (choice == 'open mail'):
            from_address = getUserInput('Please provide a from address')
            mail_to_be_opened = getMailDetailsByEmailId(from_address)
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


def inbox():
    mails = getMails()
    readMails(mails)
    postInboxMenu()


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