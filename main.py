import os.path
import stat
import sys
import time
from distutils.dir_util import copy_tree

import win32file


def locate_usb():
    drive_list = []
    driveBits = win32file.GetLogicalDrives()
    for d in range(1, 26):
        mask = 1 << d
        if driveBits & mask:
            # here if the drive is at least there
            drName = '%c:\\' % chr(ord('A') + d)
            t = win32file.GetDriveType(drName)
            if t == win32file.DRIVE_REMOVABLE:
                drive_list.append(drName)
    return drive_list


def getFileModificationDate(path):
    fileStatsObj = os.stat(path)
    seconds = fileStatsObj[stat.ST_MTIME]
    localT = time.localtime(seconds)
    timeString = time.strftime("%d %b %Y %H%M", localT)
    return timeString


def getSerialNumber(backupPath):
    serialNumberPath = os.path.join(backupPath, 'SerialNumber.txt')
    serialNumber = None
    if os.path.isfile(serialNumberPath):
        with open(serialNumberPath) as serialNumberFile:
            serialNumber = serialNumberFile.read()
    return serialNumber


def GetBackup():
    drive_info = locate_usb()
    if len(drive_info) == 0:
        print("Error: No USB drive detected")
        sys.exit(1)
    elif len(drive_info) > 1:
        print(f'Error: More than one USB drive detected - ({drive_info}')
        sys.exit(2)
    else:
        usbDrive = drive_info[0]
        backupPath = os.path.join(usbDrive, 'Backup')
        if not os.path.isdir(backupPath):
            print(f'Error: Backup directory missing - {backupPath}')
            sys.exit(3)
        else:
            modificationDate = getFileModificationDate(backupPath)
            serialNumber = getSerialNumber(backupPath)
            if serialNumber is None:
                print("Error: Serial Number file missing from backup folder")
                sys.exit(4)
            else:
                newBackupName = f'Backup from F2-{serialNumber} {modificationDate}'
                workingDir = os.getcwd()
                newPath = os.path.join(workingDir, newBackupName)
                copy_tree(backupPath, newPath)
                print(f'Success: Copied "{backupPath}" to "{newPath}"')


if __name__ == '__main__':
    GetBakup()
