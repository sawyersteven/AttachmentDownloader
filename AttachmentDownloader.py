# # # # # # # # # # # # # # # # # # # #
# User-editable variables # # # # # # #

# Time in seconds to pause between email checks
CHECKRATE = 60
# Email address to read
ADDRESS = ''
# Email account password. Leave blank to be asked for password when starting
PASSWORD = ''
# Hide password when entering in prompt
HIDEPASSWORD = True
# Email provider's imap server address
IMAPSEVER = 'imap.gmail.com'
# Mailbox to check for attachments
MAILBOX = 'Inbox'

# # # # # # # # # # # # # # # # # # # #
# # # # # # # # # # # # # # # # # # # #

import email                            # nopep8
import os                               # nopep8
import base64                           # nopep8
import imaplib                          # nopep8
import re                               # nopep8
import getpass                          # nopep8
import time                             # nopep8
import usb.core                         # nopep8
import email                            # nopep8
import tempfile                         # nopep8
import zipfile                          # nopep8
from playsound import playsound         # nopep8

MAILSERVER = None
TEMPDIR = tempfile.gettempdir()

mailRegex = r'^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$'

ALERTSOUND = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'alert.mp3')


def validateAddress(address):
    return re.search(mailRegex, address) != None


def getAddress():
    global ADDRESS

    while ADDRESS == '':
        address = input('Enter Address:')
        if validateAddress(address):
            ADDRESS = address
        else:
            m = '{} may not be a valid email adress. Use anyway? [y/n]'.format(address)
            if input(m) == 'y':
                ADDRESS = address


def getPassword():
    global PASSWORD
    PASSWORD = getpass.getpass()


def getMailBox():
    global MAILSERVER
    try:
        MAILSERVER = imaplib.IMAP4_SSL(IMAPSEVER)
        MAILSERVER.login(ADDRESS, PASSWORD)
        MAILSERVER.select(readonly=0)
    except Exception as e:
        MAILSERVER = None
        print(e)
        print('\n\n')


def main():

    if ADDRESS == '':
        getAddress()

    if PASSWORD == '':
        getPassword()

    while True:
        try:
            mainLoop()
        except Exception as e:
            print(e)
            print('\n\n')

        time.sleep(CHECKRATE)


def getDrive():
    # todo: test this
    devs = list(usb.core.find(find_all=True, bdeviceClass=8))
    if len(devs) > 0:
        return devs[0]
    else:
        return None


def mainLoop():
    global MAILSERVER
    drive = getDrive()
    if drive is None:
        return

    if MAILSERVER is None:
        getMailBox()
    if MAILSERVER is None:
        return

    tempzipfile = None

    try:
        code, response = MAILSERVER.search(None, '(UNSEEN)')
    except Exception as e:
        print(e)
        MAILSERVER = None
        return

    if code != 'OK':
        print('Mailserver responded with ' + code)
        MAILSERVER = None
        return

    msgIds = response[0].split(b' ')
    if len(msgIds) == 0 or msgIds == [b'']:
        return

    _, raw_message = MAILSERVER.fetch(msgIds[-1], 'RFC822')

    message = email.message_from_string(raw_message[0][1].decode('utf-8'))

    for part in message.walk():
        if(part.get_content_maintype() == 'multipart') or part.get('Content-Disposition') is None:
            continue
        filename = part.get_filename()
        if not filename:
            continue

        tempzipfile = os.path.join(TEMPDIR, filename)
        if os.path.isfile(tempzipfile):
            os.remove(tempzipfile)

        with open(tempzipfile, 'wb') as f:
            f.write(part.get_payload(decode=True))

        break

    if tempzipfile is None:
        return

    with zipfile.ZipFile(tempzipfile, 'r') as zipf:
        zipf.extractall(drive)

    MAILSERVER.store(msgIds[-1], '+FLAGS', '(\Seen)')
    playsound(ALERTSOUND)


if __name__ == '__main__':
    main()
