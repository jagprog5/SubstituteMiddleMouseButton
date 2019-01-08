import pythoncom, pyHook, win32ui, pyautogui

# pyHook must be manually added!
# https://www.lfd.uci.edu/~gohlke/pythonlibs/#pyhook
# Download the file corresponding to the python version being used (I'm using 3.6.6)
# pip install the file

exit_prompt_given = False
middle_btn_state = False
l_ctrl_state = False
outside_key_count = 0

def OnKeyboardEvent(event):
    if exit_prompt_given or event.WindowName is None: # occurs when program is closing
        return True

    global outside_key_count
    # only work with key inputs if solid edge is being used
    if event.WindowName.startswith("Solid Edge"):
        outside_key_count = 0 # reset outside key count once solid edge is being used
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
                        exit_with_prompt("")
                    pyautogui.mouseDown(button='middle')
                else:
                    pyautogui.mouseUp(button='middle')
        elif event.Key == "Lcontrol":
            l_ctrl_state = event.MessageName == "key down"
            if l_ctrl_state and middle_btn_state:
                exit_with_prompt("")
    else:
        outside_key_count += 1
        # increment key count if solid edge is not in focus
        if outside_key_count == 50:
            exit_with_prompt("You may have accidentally forgotten this program.\n")

    return True


def exit_with_prompt(prefix):
    global exit_prompt_given
    exit_prompt_given = True
    win32ui.MessageBox(prefix + "The program is exiting.", "Mouse Util", 4096)
    exit()

hm = pyHook.HookManager()
# get both key up and key down events
hm.KeyDown = OnKeyboardEvent
hm.KeyUp = OnKeyboardEvent
hm.HookKeyboard()
pythoncom.PumpMessages()
