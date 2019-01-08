import pythoncom
import pyHook
import pyautogui
import win32ui
import threading
import time
import win32gui
import win32com.client
import sys

# pyHook must be manually added!
# https://www.lfd.uci.edu/~gohlke/pythonlibs/#pyhook
# Download the file corresponding to the python version being used (I'm using 3.6.6)
# pip install the file

exit_prompt_given = False
middle_btn_state = False
l_ctrl_state = False
outside_key_count = 0

def OnKeyboardEvent(event):
    global exit_prompt_given
    if exit_prompt_given or event.WindowName is None:  # occurs when program is closing
        return True

    global outside_key_count
    global l_ctrl_state
    global middle_btn_state

    # Oem_3 is the '`' key, or lowercase '~'. Key near top left of keyboard
    if event.Key == "Oem_3":
        new_state = event.MessageName == "key down"
        if new_state != middle_btn_state:
            # the button state has been toggled!
            middle_btn_state = new_state
            if middle_btn_state and l_ctrl_state:
                # no need to release middle mouse button before exiting
                # the program is being exited before the button is pressed down
                exit_with_prompt("")

            # only work if solid edge is being used
            if event.WindowName.startswith("Solid Edge"):
                # change middle mouse button state based on keyboard button state
                if middle_btn_state:
                    pyautogui.mouseDown(button='middle')
                else:
                    pyautogui.mouseUp(button='middle')
    elif event.Key == "Lcontrol":
        l_ctrl_state = event.MessageName == "key down"
        if l_ctrl_state and middle_btn_state:
            # release button before exiting.
            pyautogui.mouseUp(button='middle')
            exit_with_prompt("")

    if not event.WindowName.startswith("Solid Edge"):
        outside_key_count += 1
        # increment key count if solid edge is not in focus
        if outside_key_count == 50:
            exit_with_prompt("You may have accidentally forgotten that this program is running.\n")
    else :
        outside_key_count = 0  # reset outside key count once solid edge is being used

    return True


"""
In order:
-Makes all other key inputs ignored (by setting exit_prompt_given)
-Open dialog box in thread
-While thread is running, elevate the dialog
-Join. Wait for dialog to close
-Exit
"""
def exit_with_prompt(prefix):
    global exit_prompt_given
    exit_prompt_given = True
    dialog_title = "Mouse Util"
    t = threading.Thread(target=_message_thread, args=(dialog_title, prefix + "The program is exiting."))
    t.start()
    time.sleep(0.1)
    win32gui.EnumWindows(_prompt_enum_handle, dialog_title)
    t.join()
    sys.exit()


# run within thread, because message box locks thread until it is closed
def _message_thread(title, text):
    win32ui.MessageBox(text, title)


def _prompt_enum_handle(hwnd, title):
    if win32gui.GetWindowText(hwnd) == title and win32gui.GetClassName(hwnd) == "#32770":
        # https://stackoverflow.com/a/30314197/5458478
        # ¯\_(ツ)_/¯
        win32com.client.Dispatch("WScript.Shell").SendKeys('%')
        win32gui.SetForegroundWindow(hwnd)


hm = pyHook.HookManager()
# get both key up and key down events
hm.KeyDown = OnKeyboardEvent
hm.KeyUp = OnKeyboardEvent
hm.HookKeyboard()
pythoncom.PumpMessages()
