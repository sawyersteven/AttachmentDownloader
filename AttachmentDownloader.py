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
import zipfile                          # nopep8
from playsound import playsound         # nopep8
from platform import system             # nopep8

SYSTEM = system()
if SYSTEM == 'Windows':
    import ctypes
else:
    import pyudev

MAILSERVER = None
os.chdir(os.path.dirname(os.path.realpath(__file__)))
TEMPDIR = os.path.abspath('./TEMP')
ALERTSOUND = './alert.mp3'
GETDRIVE = None

mailRegex = r'^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$'


def _getDriveWindows():
    drive_letters = []
    n_drives = ctypes.windll.kernel32.GetLogicalDrives()
    for i in range(0, 25):
        j = 2**i
        if n_drives & j > 0:
            drive_letters.append((chr(65 + i) + ':/').encode())

    for i in drive_letters:
        if ctypes.windll.kernel32.GetDriveTypeA(i) == 2:
            print('Using external drive {}'.format(i))
            return i.decode('utf-8')
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


def downloadAttachment(msgid):
    _, raw_message = MAILSERVER.fetch(msgid, 'RFC822')

    message = email.message_from_string(raw_message[0][1].decode('utf-8'))

    for part in message.walk():
        if(part.get_content_maintype() == 'multipart') or part.get('Content-Disposition') is None:
            continue
        filename = part.get_filename()
        if not filename:
            continue

        tempfile = os.path.join(TEMPDIR, '{}_{}'.format(msgid, filename))
        if os.path.isfile(tempfile):
            print('Removing old zip file from temp dir')
            os.remove(tempfile)

        print('Downloading zip to local temp dir as {}'.format(filename))
        with open(tempfile, 'wb') as f:
            f.write(part.get_payload(decode=True))

        break

    MAILSERVER.store(msgid, '+FLAGS', r'(\Seen)')
    return tempfile


def extractZip(file, dst):
    print('Extracting zip to {}'.format(dst))

    with zipfile.ZipFile(file, 'r') as zipf:
        zipf.extractall(dst)


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

    print('Found usb drive at {}'.format(drive))

    if MAILSERVER is None:
        getMailBox()
    if MAILSERVER is None:
        print('Could not connect to email service')
        return

    try:
        code, response = MAILSERVER.search(None, '(UNSEEN)')
    except Exception as e:
        print(e)
        MAILSERVER = None
        return

    if code != 'OK':
        print('Mailserver responded with {}'.format(code))
        MAILSERVER = None
        return

    MAILSERVER.select(readonly=0)
    msgIds = response[0].split(b' ')
    if len(msgIds) == 0 or msgIds == [b'']:
        print('No unread emails on server')
        return

    attachments = []
    for msgid in msgIds[::-1]:
        attachments.append(downloadAttachment(msgid))

    for filename in attachments:
        extractZip(filename, drive)
        os.remove(filename)

    playsound(ALERTSOUND)


if __name__ == '__main__':
    main()
