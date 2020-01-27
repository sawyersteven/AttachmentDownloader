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
import email                            # nopep8
import tempfile                         # nopep8
import zipfile                          # nopep8
from playsound import playsound         # nopep8
from platform import system             # nopep8

SYSTEM = system()
if SYSTEM == 'Windows':
    import ctypes
else:
    import pyudev

MAILSERVER = None
TEMPDIR = tempfile.gettempdir()
os.chdir(os.path.dirname(os.path.realpath(__file__)))
ALERTSOUND = './alert.mp3'
GETDRIVE = None

mailRegex = r'^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$'


def _getDriveWindows():
    drive_letters = []
    n_drives = ctypes.windll.kernel32.GetLogicalDrives()
    for i in range(0, 25):
        j = 2**i
        if n_drives & j > 0:
            drive_letters.append((chr(65 + i) + ":/").encode())

    for i in drive_letters:
        if ctypes.windll.kernel32.GetDriveTypeA(i) == 2:
            print("Using external drive " + i)
            return i
    return None


def _getDriveNix():
    ctx = pyudev.Context()
    usb_drives = list(ctx.list_devices(ID_BUS='usb', subsystem='block'))

    devnames = []
    for drive in usb_drives:
        if drive.get('DEVTYPE') == 'partition' and drive.get('DEVNAME'):
            devnames.append(drive.get('DEVNAME'))

    if len(devnames) == 0:
        return None

    mounts = {}
    with open('/proc/mounts') as f:
        for line in f.readlines():
            parts = line.split(' ')
            mounts[parts[0]] = parts[1]

    for dev in devnames:
        p = mounts.get(dev)
        if p != None:
            return p

    return None


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
    global GETDRIVE

    SYSTEM = system()
    if SYSTEM == 'Windows':
        GETDRIVE = _getDriveWindows
    else:
        GETDRIVE = _getDriveNix

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


def mainLoop():
    global MAILSERVER

    print('\n\n')
    print('Checking {} for new attachments'.format(ADDRESS))

    drive = GETDRIVE()
    if drive is None:
        print('USB drive not found, skipping cycle')
        return

    print('Found usb drive at ' + drive)

    if MAILSERVER is None:
        getMailBox()
    if MAILSERVER is None:
        print('Could not connect to email service')
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
        print('No unread emails on server')
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
            print('Removing existing zip file')
            os.remove(tempzipfile)

        print('Downloading zip to local temp dir')
        with open(tempzipfile, 'wb') as f:
            f.write(part.get_payload(decode=True))

        break

    if tempzipfile is None:
        return

    print('Extracting zip to drive: ' + drive)

    with zipfile.ZipFile(tempzipfile, 'r') as zipf:
        zipf.extractall(drive)

    MAILSERVER.store(msgIds[-1], '+FLAGS', '(\Seen)')
    playsound(ALERTSOUND)


if __name__ == '__main__':
    main()
