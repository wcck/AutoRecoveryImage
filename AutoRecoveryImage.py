import os
import string
import ctypes, sys
import locale
import win32con
import win32api
import pyautogui
import time
import logging
from ctypes import windll
from win32 import win32gui
from win32 import win32clipboard as w

# import win32process

# Log
logging.basicConfig(filename='AutoRecoveryImage.log',
                    filemode='a',
                    format='%(asctime)s,%(msecs)d %(name)s %(levelname)s %(message)s',
                    datefmt='%H:%M:%S',
                    level=logging.DEBUG)


# Check whether insert USB on NB/PC
def get_drives():
    drives = []
    bitmask = windll.kernel32.GetLogicalDrives()
    for letter in string.ascii_uppercase:
        if bitmask & 1:
            drives.append(letter)
        bitmask >>= 1
    return drives

#run as administrator for recovery image tool
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()        
    except:
        return False

def left_click(left, top, right, bottom, hwnd):
    # Calculate coordinate for event button
    x = int(left+(right-left)/2)
    y = int(top+(bottom-top)/2)  
    logging.info("coordinate(%d, %d)" %(x, y))   
    print("coordinate(%d, %d)" %(x, y))
    pyautogui.moveTo(x, y)    

    # Click  
    win32api.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON)
    time.sleep(0.2)
    win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON)
    print('Left Click')

def judge_locale():
    # Get locale
    locLang = locale.getdefaultlocale()
    print(locLang[0])
    return locLang[0]

def PASS():
    logging.info("Recovery Drive Done!")
    return True    

def FAIL():
    sys.exit()    
    return False    

# if is_admin() : 
if True :
    driveNumber = set(get_drives())
    print(driveNumber)
    if len(driveNumber) < 2:
        logging.error("[ERROR]Need to insert USB on platform.")
        FAIL()
    
    # Disable UAC
    # Code of your program here
    RecoveryDrivePath = os.path.join(os.environ.get('windir'), 'system32','RecoveryDrive.exe')
    if RecoveryDrivePath == "":
        logging.error("[ERROR]Cannot find tool path.")
        FAIL()
    logging.info("Tool Path : %s" %(RecoveryDrivePath))
    logging.info("Execute Recovery Drive")
    os.startfile(RecoveryDrivePath)

    logging.info("Waiting for 1 seconds")
    time.sleep(1)

    # Judge english or chinese on environment
    localeType = judge_locale()
    titleHwnd = 0
    if localeType == "zh_TW" :
        titleHwnd = win32gui.FindWindow("NativeHWNDHost", u"修復磁碟機")
    elif localeType == "en-US" :
        titleHwnd = win32gui.FindWindow("NativeHWNDHost", "Recovery Drive")
    # else:
    #     titleHwnd = win32gui.FindWindow("NativeHWNDHost", "Recovery Drive")

    # Get recovery drive tool title
    title = win32gui.GetWindowText(titleHwnd)
    logging.info("title %s" %(title))

    print(titleHwnd)
    ClassName = win32gui.GetClassName(titleHwnd)
    print("%s" %(ClassName))

    # Trigger button events
    child = []   
    def all_ok(hwnd, parm):
        child.append(hwnd)
        ClassName = win32gui.GetClassName(hwnd)
        title = win32gui.GetWindowText(hwnd)
        print("title : %s, class : %s, hwnd : %d" %(title, ClassName, hwnd))
        
        if title == u"下一步(&N)" :
            print(hwnd)
            left,top,right,bottom = win32gui.GetWindowRect(hwnd)
            # Click mouse left event
            left_click(left, top, right, bottom, hwnd)
            child.clear()
        
        elif title == u"下一步" :
            print(hwnd)
            left,top,right,bottom = win32gui.GetWindowRect(hwnd)
            # Click mouse left event
            left_click(left, top, right, bottom, hwnd)
            child.clear()

        elif title == "Next" :
            print(hwnd)
            left,top,right,bottom = win32gui.GetWindowRect(hwnd)
            # Click mouse left event
            left_click(left, top, right, bottom, hwnd)
            child.clear()

        elif title == u"建立" :
            print(hwnd)
            left,top,right,bottom = win32gui.GetWindowRect(hwnd)
            # Click mouse left event
            left_click(left, top, right, bottom, hwnd)
            child.clear()

        elif title == "Create" :
            print(hwnd)
            left,top,right,bottom = win32gui.GetWindowRect(hwnd)
            # Click mouse left event
            left_click(left, top, right, bottom, hwnd)
            child.clear()
        
        # elif title == u"完成(&F)" :
        #     print(hwnd)
        #     left,top,right,bottom = win32gui.GetWindowRect(hwnd)
        #     # Click mouse left event
        #     left_click(left, top, right, bottom, hwnd)
        #     child.clear()
        #     PASS()
        
        elif title == u"完成" :
            print(hwnd)
            left,top,right,bottom = win32gui.GetWindowRect(hwnd)
            # Click mouse left event
            left_click(left, top, right, bottom, hwnd)
            child.clear()
            PASS()

        elif title == "Finish" :
            print(hwnd)
            left,top,right,bottom = win32gui.GetWindowRect(hwnd)
            # Click mouse left event
            left_click(left, top, right, bottom, hwnd)
            child.clear()
            PASS()
        
        child.clear()
    
    # Polling unitl show button events
    while(True):
        win32gui.EnumChildWindows(titleHwnd, all_ok, None)
        # # Interrupt Tool
        # titleHwnd = 0
        # if localeType == "zh_TW" :
        #     titleHwnd = win32gui.FindWindow("NativeHWNDHost", u"修復磁碟機")
        # elif localeType == "en-US" :
        #     titleHwnd = win32gui.FindWindow("NativeHWNDHost", "Recovery Drive")        
        

        # # Get recovery drive tool title
        # title = win32gui.GetWindowText(titleHwnd)
        # logging.info("title %s" %(title))

        # print(titleHwnd)
        # if titleHwnd == 0 :
        #     logging.error("[Error]Interrupt by user.")
        #     FAIL()

        # Polling once every 10 seconds
        time.sleep(10)
        print(child)    
    
    # os.system("pause")
    # # Execute next step

    
else :
    # Re-run the program with admin rights
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
