# -*- coding: utf-8 -*-
"""
Created on Tue Dec  7 06:40:56 2021

@author: harim
"""

####################### IMPORTS AND DICTIONARIES ##############################

from pynput.mouse import Listener as ML
from pynput.keyboard import Listener as KL, Key
from pyautogui import click, press, hotkey, scroll, keyDown, keyUp
from pyperclip import copy, paste
from time import sleep

# used to simply hit the key instead of pressing down and then up
single_keys = ['\t', '\n', '\r', ' ', '!', '"', '#', '$', '%', '&', "'", '(',
    ')', '*', '+', ',', '-', '.', '/', '0', '1', '2', '3', '4', '5', '6', '7',
    '8', '9', ':', ';', '<', '=', '>', '?', '@', '[', '\\', ']', '^', '_', '`',
    'a', 'b', 'c', 'd', 'e','f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o',
    'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z', '{', '|', '}', '~']

# used to translate common hotkeys to the format used by pyautogui
translate_hotkey = {'\\x01' : ['ctrl', 'a'],
                    '\\x02' : ['ctrl', 'b'],
                    '\\x03' : ['ctrl', 'c'],
                    '\\x05' : ['ctrl', 'e'],
                    '\\x04' : ['ctrl', 'd'],
                    '\\x06' : ['ctrl', 'f'],
                    '\\x07' : ['ctrl', 'g'],
                    '\\x08' : ['ctrl', 'h'],
                    '\\x09' : ['ctrl', 'i'],
                    '\\x0a' : ['ctrl', 'j'],
                    '\\x0b' : ['ctrl', 'k'],
                    '\\x0c' : ['ctrl', 'l'],
                    '\\x0d' : ['ctrl', 'm'],
                    '\\x0e' : ['ctrl', 'n'],
                    '\\x0f' : ['ctrl', 'o'],
                    '\\x10' : ['ctrl', 'p'],
                    '\\x11' : ['ctrl', 'q'],
                    '\\x12' : ['ctrl', 'r'],
                    '\\x13' : ['ctrl', 's'],
                    '\\x14' : ['ctrl', 't'],
                    '\\x15' : ['ctrl', 'u'],
                    '\\x16' : ['ctrl', 'v'],
                    '\\x17' : ['ctrl', 'w'],
                    '\\x18' : ['ctrl', 'x'],
                    '\\x19' : ['ctrl', 'y'],
                    '\\x1a' : ['ctrl', 'z']}

# used to translate special key names to the ones used by pyautogui
translate_key = {"ctrl_l":"ctrl",
                 "ctrl_r":"ctrl",
                 "ctrl":"ctrl",
                 "caps_lock":"capslock",
                 "alt":"alt",
                 "alt_l":"alt",
                 "alt_r":"altright",
                 "alt_gr":"altright",
                 "page_up":"pageup",
                 "page_down":"pagedown",
                 "cmd":"win",
                 "print_screen":"printscreen",
                 "media_previous":"prevtrack",
                 "media_play_pause":"playpause",
                 "media_next":"nexttrack",
                 "shift":"shift",
                 "shift_l":"shift",
                 "shift_r":"shift",
                 "media_volume_mute":"volumemute",
                 "media_volume_down":"volumedown",
                 "media_volume_up":"volumeup",
                 "\\\\":"\\",
                 "\"\"":"'"}

# used to release all the keys before a fail-safe
additional_keys = ['accept', 'add', 'alt', 'altleft', 'altright', 'apps', 'backspace',
    'browserback', 'browserfavorites', 'browserforward', 'browserhome',
    'browserrefresh', 'browsersearch', 'browserstop', 'capslock', 'clear',
    'convert', 'ctrl', 'ctrlleft', 'ctrlright', 'decimal', 'del', 'delete',
    'divide', 'down', 'end', 'enter', 'esc', 'escape', 'execute', 'f1', 'f10',
    'f11', 'f12', 'f13', 'f14', 'f15', 'f16', 'f17', 'f18', 'f19', 'f2', 'f20',
    'f21', 'f22', 'f23', 'f24', 'f3', 'f4', 'f5', 'f6', 'f7', 'f8', 'f9',
    'final', 'fn', 'hanguel', 'hangul', 'hanja', 'help', 'home', 'insert', 'junja',
    'kana', 'kanji', 'launchapp1', 'launchapp2', 'launchmail',
    'launchmediaselect', 'left', 'modechange', 'multiply', 'nexttrack',
    'nonconvert', 'num0', 'num1', 'num2', 'num3', 'num4', 'num5', 'num6',
    'num7', 'num8', 'num9', 'numlock', 'pagedown', 'pageup', 'pause', 'pgdn',
    'pgup', 'playpause', 'prevtrack', 'print', 'printscreen', 'prntscrn',
    'prtsc', 'prtscr', 'return', 'right', 'scrolllock', 'select', 'separator',
    'shift', 'shiftleft', 'shiftright', 'sleep', 'space', 'stop', 'subtract', 'tab',
    'up', 'volumedown', 'volumemute', 'volumeup', 'win', 'winleft', 'winright', 'yen',
    'command', 'option', 'optionleft', 'optionright']

############################### FUNCTIONS #####################################

# for storing all the mouse clicks and keyboard strokes
fail_safe_log = []
log = []

# the following logging functions are used later by the pynput listener
def on_click(x, y, button, pressed):
    if pressed:
        log.append([x, y, str(button)[7:]])
        # print('Mouse clicked at ({0}, {1}) with {2}'.format(x, y, button))

def on_scroll(x, y, dx, dy):
    log.append([x, y, dx, dy])
    # print('Mouse scrolled at ({0}, {1})({2}, {3})'.format(x, y, dx, dy))
    
def on_press(key):
    if "Key." in str(key):        
        log.append([str(key)[4:].replace("'", ""), "down"])
        # print('  {0} key pressed'.format(key).replace("Key.", ""))
    else:
        log.append([str(key).replace("'", ""), "down"]) 
        # print('  {0} key pressed'.format(key))
    if key == Key.esc:
        return False

def on_release(key):
    if "Key." in str(key):        
        log.append([str(key)[4:], "up"])
        # print('  {0} key released'.format(key).replace("Key.", ""))
    else:
        log.append([str(key).replace("'", ""), "up"]) 
        # print('  {0} key released'.format(key))

def fail_safe(key):
    '''Stores the pressed keys in the fail_safe_log for later stop of the execution.'''
    fail_safe_log.append(key)
    if key == Key.esc:
        return False

def timer(wait_time, message):
    print("\n"+message, end="")
    for sec in range(wait_time):
        sleep(0.25)
        print(wait_time-sec, end="")
        sleep(0.25)
        print(".", end="")
        sleep(0.25)
        print(".", end="")
        sleep(0.25)
        print(". ", end="")
    print()

def warning_message(before_recording=True):
    if before_recording:
        print("\n  ******* IMPORTANT: Please read carefully! ********\n")
        print("   - To stop recording, press 'esc'.")
        print("   - To stop the execution, press 'esc' or quickly")
        print("     move the mouse to a corner of the screen.\n")
        print("  WARNING: The actions performed are not reversible.")
        print("  If anything goes wrong, press 'esc' repeteadly.\n")
        print("  **************************************************\n")
    else:
        print("\n  ************* IMPORTANT REMINDER **************\n")
        print("   - To stop the execution, press 'esc' or quickly")
        print("     move the mouse to a corner of the screen.")
        print("   - The actions performed are not reversible. If")
        print("     anything goes wrong, press 'esc' repeteadly.\n")
        print("  **************************************************\n")

def validate_input(options_menu):
    '''Displays a message requesting user input and accepts only integers.'''
    while True:
        print(options_menu)
        number = input("\n  >> ")
        if number == "":
            number = 0
            break
        try:
            number = int(number)
            break
        except:
            print("\n  ERROR: Received '{}' as input.".format(number),
                  "\n  Expected an integer (0,1,2,...).\n")
            pass
    return number

def record_instructions(waiting_time, verbose=False):    
    '''Starts the listener after a given window of time. Captures all the mouse
       scrolls and clicks as wells as the keyboard strokes in a log.'''
    
    if waiting_time > 0:
        timer(waiting_time, "  Start recording in " + str(waiting_time) + " seconds: ")
    else: 
        sleep(0.2)
    
    print("\n\t     * * * NOW RECORDING * * *")    
    mouse_listener = ML(on_click=on_click, on_scroll=on_scroll)
    mouse_listener.start()
    with KL(on_press=on_press, on_release=on_release) as keyboard_listener:
        keyboard_listener.join()
        log.pop()
    mouse_listener.stop()
    
    if verbose and not log == []: 
        print("\n  RECORDED INTRUCTIONS:\n")
        [ print(x) for x in log ]
        
    return log

def execute_instructions(instructions_list, verbose=False):
    '''Takes a log with recorded instructions in pynput format and
       translates them into pyautogui format for execution.'''
    
    if instructions_list == []:
        return print("\n  INPUT EXCEPTION: No instructions recorded for execution.")

    if verbose: print("\n  EXECUTION LOG:\n")
    
    # Start the listener to be able to fail-safe the executions with 'esc'
    KB_listener = KL(on_press=fail_safe)
    KB_listener.start()
    fail_safe_presses = []
    
    contador = 1
    for i in instructions_list:
        sleep(0.1)
        
        # If 'esc' key is pressed, stop the execution and exit
        if Key.esc in fail_safe_log:
            KB_listener.stop()
            if not fail_safe_presses == []: # realease all keys
                for k in fail_safe_presses:
                    keyUp(k)
                print("\n  RELEASED KEYS: " + str(fail_safe_presses))
            print("\n  FAIL-SAFE: fail-safe was triggered from pressing the 'esc' key."
                  "\n  The intructions's execution was terminated successfully at [{}]/[{}].\n".format(contador-1, len(instructions_list)))
            break
        
        # check if the intruction is a key stroke
        if len(i) == 2:
            
            # check if the key is a character
            if i[0] in single_keys and i[1] == "down":
                if i[0] not in fail_safe_presses: fail_safe_presses.append(i[0])
                press(i[0])
                if verbose: print("  [{}] Hit {} key".format(contador, i[0]))
                
            # check if it's a hotkey combination
            elif i[0] in translate_hotkey.keys() and i[1] == "down":
                hotkey(translate_hotkey[i[0]][0], translate_hotkey[i[0]][1])
                if verbose: print("  [{}] Send hotkey {} + {}".format(contador, translate_hotkey[i[0]][0], translate_hotkey[i[0]][1]))
            
            # check if it's an special key
            elif i[0] in translate_key.keys():
                if i[1] == "down":
                    if translate_key[i[0]] not in fail_safe_presses: fail_safe_presses.append(translate_key[i[0]])
                    keyDown(translate_key[i[0]])
                    if verbose: print("  [{}] Press down {} key".format(contador, translate_key[i[0]]))
                else:
                    keyUp(translate_key[i[0]])
                    if verbose: print("  [{}] Press up {} key".format(contador, translate_key[i[0]]))
            
            else: # try-except block for the remaining keys
                try:
                    if i[1] == "down":
                        if i[0] not in fail_safe_presses: fail_safe_presses.append(i[0])
                        press(i[0])
                        if verbose: print("  [{}] Hit {} key".format(contador, i[0]))
                except:
                    print("\n  ERROR: Invalid key stroke {} (instruction [{}] skipped).\n".format(i[0], contador))

        # check if the intruction is a click                 
        elif len(i) == 3:
            click(i[0], i[1], button=i[2], duration=0.15)
            if verbose: print('  [{}] Mouse clicked ({}, {}) with {} button'.format(contador, i[0], i[1], i[2]))
        
        else: # scroll the mouse
            scroll(200*i[3], x=i[0], y=i[1])
            if verbose: print('  [{}] Mouse scrolled at ({}, {}) ({}, {})'.format(contador, i[0], i[1], i[2], i[3]))
        
        contador += 1

def str_to_list(string_ls):
    '''Takes a string that corresponds to the list generated by the
       record_instructions function and turns it into a list again.'''        
    try:
        # creates a list of strings where each one is an instruction
        string_ls = string_ls[2:-2].split('], [')
        
        # for each instruction, split the info and store each its values
        results = []
        [ results.append(ins.split(', ')) for ins in string_ls ]
        
        for res in results:
            
            # remove the commas to keep the string as it is
            if len(res) == 2: 
                results[results.index(res)] = [ x[1:-1] for x in res ]
            
            # takes [a, b, c] and returns [int(a), int(b), str(c[1:-1])]
            elif len(res) == 3:
                results[results.index(res)][0] = int(res[0])
                results[results.index(res)][1] = int(res[1])
                results[results.index(res)][2] = str(res[2][1:-1])
            
            # converts each item of the list into an integer
            elif len(res) == 4:
                results[results.index(res)] = [ int(x) for x in res ]
            
            else:
                return []
            
        return results
    
    except:
        return []    

################################# MAIN #########################################

exit_program = False

warning_message(before_recording=True)

while True:
    
    menu_rec = "   1 --> Start recording right away.\n" + "   2 --> Set the timer before recording.\n" + "   3 --> Use a recording from clipboard.\n\n" + "  Enter an option or press 'enter' to exit the program."
    rec_comm = validate_input(menu_rec)
    
    if rec_comm == 0:
        exit_program = True
        print("\n\t   * * * CANCEL AND EXIT * * *\n")
        break
    
    # start recording intructions right away
    elif rec_comm == 1:
        inst = record_instructions(0) # to debug set verbose=True
        break
    
    # set the timer for the recording
    elif rec_comm == 2:        
        prompt = "\n  Enter the number of seconds for the timer."
        wait = validate_input(prompt)
        inst = record_instructions(wait) # to debug set verbose=True
        break
    
    # use the instructions from the clipboard
    elif rec_comm == 3:
        inst = str_to_list(paste())
        if inst == []:
            print("\n  ERROR: The instructions found in the clipboard were",
                  "\n  invalid or incomplete. Please select another option. \n")
            continue
        break
        
    else: # No option assinged to the input received
        print("\n  ERROR: Received '{}' as input.".format(rec_comm),
              "\n  Expected '1', '2', '3' or '0' ('enter').\n")


if not exit_program: warning_message(before_recording=False)

if not exit_program:
    
    while True:
        menu = "  What do you want to do next with the recording?\n\n" + "   1 --> Execute the recording once.\n" + "   2 --> Loop the execution x number of times.\n" + "   3 --> Save recording to the clipboard.\n\n" + "  To cancel and exit, press 'enter'."
        exec_comm = validate_input(menu)
        
        if exec_comm == 0:
            print("\n\t   * * * CANCEL AND EXIT * * *\n")
            break
        
        # execute instrucions once
        elif exec_comm == 1:
            timer(3, "  Starting execution in 3 seconds: ")
            execute_instructions(inst, verbose=True) # to debug set verbose=True
            inst = str(inst)
            copy(inst)
            print("\n\t   * * * END OF EXECUTION * * *")
            break
        
        # execute instrucions in a loop
        elif exec_comm == 2:
            
            print()
            prompt = "  Enter the number of iterations for the loop."
            iterations = validate_input(prompt)
        
            timer(3, "  Starting execution in 3 seconds: ")
            print()
            for i in range(iterations):
                print("\n  Iteration number {}/{} in progress:".format(i+1,iterations))
                execute_instructions(inst, verbose=True) # to debug set verbose=True
            
            print("\n\t   * * * END OF EXECUTION * * *")            
            break
        
        # save instructions to the clipboard
        elif exec_comm == 3:
            copy(str(inst))
            print("\n     * * * RECORDING SAVED TO CLIPBOARD * * *")            
            break
            
        else: # No option assigned to the input received
            print("\n  ERROR: Received '{}' as input.".format(exec_comm),
                  "\n  Expected '1', '2' or '0' ('enter').\n")
    
input()
