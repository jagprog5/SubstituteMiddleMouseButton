import pythoncom, pyHook, win32ui, pyautogui

# pyHook must be manually added!
# https://www.lfd.uci.edu/~gohlke/pythonlibs/#pyhook
# Download the file corresponding to the python version being used (I'm using 3.6.6)
# pip install the file

middle_btn_state = False
l_ctrl_state = False

def OnKeyboardEvent(event):
    # only work with key inputs if solid edge is being used
    # event is none if the program is about to close
    if event.WindowName is not None and event.WindowName.startswith("Solid Edge"):
        global l_ctrl_state
        # Oem_3 is the '`' key, or lowercase '~'. Key near top left of keyboard
        if event.Key == "Oem_3":
            new_state = event.MessageName == "key down"
            global middle_btn_state
            if new_state != middle_btn_state:
                # the button state has been toggled!
                middle_btn_state = new_state
                if middle_btn_state:
                    if l_ctrl_state:
                        # 4096 to bring project to top
                        win32ui.MessageBox("The program is exiting.", "", 4096)
                        exit()
                    pyautogui.mouseDown(button='middle')
                else:
                    pyautogui.mouseUp(button='middle')
        elif event.Key == "Lcontrol":
            l_ctrl_state = event.MessageName == "key down"
    return True


hm = pyHook.HookManager()
# get both key up and key down events
hm.KeyDown = OnKeyboardEvent
hm.KeyUp = OnKeyboardEvent
hm.HookKeyboard()
pythoncom.PumpMessages()
