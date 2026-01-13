#!/usr/bin/env python 
# -*- coding: utf-8 -*- 
 
# Quick script to monitor outlook folder for new tickets
# Displays big numbers on screen so we don't miss them.
# press ENTER to mark everything read.
 
import time 
import random 
import os 
import sys 

# --- CONFIG ---
# Put the folder name here
TARGET_FOLDER = "YOUR_TARGET_FOLDER_NAME_HERE" 
# ----------------

# Check for libraries
try: 
    import win32com.client 
except ImportError: 
    print("Missing pywin32. pip install pywin32")
    sys.exit(1) 

try: 
    import msvcrt # windows only
except ImportError: 
    print("Windows only script.")
    sys.exit(1) 
 
# Try to load colorama for green text, otherwise just white
try: 
    from colorama import init, Fore, Style 
    init() 
    COLORS = True 
except ImportError: 
    COLORS = False 
 
MATRIX_WIDTH = 60 # just a guess for screen width
 
# ASCII digits - don't touch formatting
BIG_DIGITS = { 
    '0': [ 
        "        0000000        ", 
        "      00        00      ", 
        "     00          00     ", 
        "    00            00    ", 
        "   00              00   ", 
        "  00                00  ", 
        "  00                00  ", 
        " 00                  00 ", 
        " 00                  00 ", 
        " 00                  00 ", 
        " 00                  00 ", 
        " 00                  00 ", 
        "  00                00  ", 
        "  00                00  ", 
        "   00              00   ", 
        "    00            00    ", 
        "     00          00     ", 
        "      00        00      ", 
        "        0000000        ", 
        "                       " 
    ], 
    '1': [ 
        "           11           ", 
        "          111           ", 
        "         1111           ", 
        "           11           ", 
        "           11           ", 
        "           11           ", 
        "           11           ", 
        "           11           ", 
        "           11           ", 
        "           11           ", 
        "           11           ", 
        "           11           ", 
        "           11           ", 
        "           11           ", 
        "           11           ", 
        "           11           ", 
        "           11           ", 
        "           11           ", 
        "           11           ", 
        "                       " 
    ], 
    '2': [ 
        "       2222222       ", 
        "     22        22    ", 
        "    22          22   ", 
        "    22          22   ", 
        "              22    ", 
        "             22     ", 
        "            22      ", 
        "           22       ", 
        "          22        ", 
        "        22          ", 
        "       22           ", 
        "      22            ", 
        "     22             ", 
        "    22              ", 
        "    22              ", 
        "    222222222222    ", 
        "                    ", 
        "                    ", 
        "                    ", 
        "                    " 
    ], 
    '3': [ 
        "      33333333       ", 
        "     33       33      ", 
        "    33         33     ", 
        "             33      ", 
        "            33      ", 
        "          333        ", 
        "           3333      ", 
        "              333    ", 
        "               33    ", 
        "                33  ", 
        "                33  ", 
        " 33             33  ", 
        "  33            33    ", 
        "   33          33     ", 
        "    33        33      ", 
        "      3333333        ", 
        "                    ", 
        "                    ", 
        "                    ", 
        "                    " 
    ], 
    '4': [ 
        "            44        ", 
        "           444        ", 
        "          44 4        ", 
        "         44  4        ", 
        "        44   4        ", 
        "       44    4        ", 
        "      44     4        ", 
        "     44      4        ", 
        "    44       4        ", 
        "   44        4        ", 
        "   444444444444444    ", 
        "              4        ", 
        "              4        ", 
        "              4        ", 
        "              4        ", 
        "                       ", 
        "                       ", 
        "                       ", 
        "                       ", 
        "                       " 
    ], 
    '5': [ 
        "  5555555555555   ", 
        "  55              ", 
        "  55              ", 
        "  55              ", 
        "  55555555555     ", 
        "             55   ", 
        "              55  ", 
        "              55  ", 
        "              55  ", 
        " 55           55  ", 
        "  55         55   ", 
        "   55        55    ", 
        "    55      55     ", 
        "     55    55      ", 
        "       555         ", 
        "                 ", 
        "                 ", 
        "                 ", 
        "                 ", 
        "                 " 
    ], 
    '6': [ 
        "        6666666      ", 
        "      66        66    ", 
        "     66          66   ", 
        "    66              ", 
        "   66                ", 
        "  66    6666666       ", 
        "  66   66       66    ", 
        "  66  66         66    ", 
        "  6666           66   ", 
        "  66             66   ", 
        "  66             66   ", 
        "   66           66    ", 
        "    66         66     ", 
        "     66       66      ", 
        "       666666        ", 
        "                     ", 
        "                     ", 
        "                     ", 
        "                     ", 
        "                     " 
    ], 
    '7': [ 
        " 777777777777777  ", 
        "            77    ", 
        "           77     ", 
        "          77      ", 
        "         77       ", 
        "        77        ", 
        "       77         ", 
        "      77          ", 
        "     77          ", 
        "    77            ", 
        "   77             ", 
        "  77              ", 
        " 77               ", 
        " 77               ", 
        " 77               ", 
        "                 ", 
        "                 ", 
        "                 ", 
        "                 ", 
        "                 " 
    ], 
    '8': [ 
        "      8888888       ", 
        "    88        88    ", 
        "   88          88    ", 
        "   88          88    ", 
        "    88        88     ", 
        "      8888888       ", 
        "    88        88    ", 
        "   88          88    ", 
        "   88          88    ", 
        "   88          88    ", 
        "   88          88    ", 
        "    88        88     ", 
        "      8888888       ", 
        "                    ", 
        "                    ", 
        "                    ", 
        "                    ", 
        "                    ", 
        "                    ", 
        "                    " 
    ], 
    '9': [ 
        "        9999999      ", 
        "      99        99    ", 
        "     99          99   ", 
        "    99          99    ", 
        "    99          99    ", 
        "     99        99     ", 
        "       9999999       ", 
        "              99     ", 
        "              99     ", 
        "              99     ", 
        "             99      ", 
        "            99       ", 
        "            99       ", 
        "                     ", 
        "                     ", 
        "                     ", 
        "                     ", 
        "                     ", 
        "                     ", 
        "                     " 
    ] 
} 
 
def stretch_line(line, factor=2): 
    return "".join(c * factor for c in line) 
 
def print_big_num(num_str): 
    # draws the number line by line
    lines = [ 
        [stretch_line(row, 2) for row in BIG_DIGITS[d]] 
        for d in num_str 
    ] 
    for row in range(20): 
        print("  ".join(art[row] for art in lines)) 

def do_matrix_rain(duration=2.0): 
    # Just makes it look cool when stuff is happening
    start = time.time() 
    while time.time() - start < duration: 
        line = "" 
        for _ in range(MATRIX_WIDTH): 
            ch = random.choice("01X|/\\{}[]()#$%^&*+;:ABC") 
            if COLORS: 
                line += Fore.GREEN + ch + Style.RESET_ALL 
            else: 
                line += ch 
        print(line) 
        time.sleep(0.05) 
    
    os.system('cls' if os.name == 'nt' else 'clear') 

def get_folder(): 
    # Grab Outlook instance
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI") 
    
    # Iterate roots to find the public folder or shared mailbox
    for i in range(1, outlook.Folders.Count + 1): 
        root = outlook.Folders.Item(i) 
        try: 
            return root.Folders.Item(TARGET_FOLDER) 
        except: 
            # keep looking
            pass 
    raise RuntimeError(f"Couldn't find folder: {TARGET_FOLDER}") 
 
def get_count(folder): 
    # restrict is faster than looping all items
    try: 
        return folder.Items.Restrict("[UnRead] = True").Count 
    except: 
        return folder.UnreadItemCount 
 
def mark_read(folder): 
    try: 
        # Grab unread items
        items = folder.Items.Restrict("[UnRead] = True") 
        # Convert to list so we don't modify the collection while looping
        todo = [i for i in items] 
        for item in todo: 
            if item.UnRead: 
                item.UnRead = False 
    except Exception as e: 
        print("Failed to mark read:", e)
        time.sleep(2) 
 
# Main execution
def run(): 
    os.system('cls' if os.name == 'nt' else 'clear') 
    last_count = None
 
    while True: 
        # Check if user hit Enter
        if msvcrt.kbhit(): 
            key = msvcrt.getwch() 
            if key == '\r': # Enter key
                try: 
                    f = get_folder()
                    mark_read(f)
                    os.system('cls' if os.name == 'nt' else 'clear')
                    print_big_num("0") 
                    last_count = 0 
                except Exception as e: 
                    print("Error:", e) 
                time.sleep(2) 
                continue 
 
        # Update loop
        try: 
            f = get_folder()
            count = get_count(f) 
        except Exception as e: 
            print("Outlook error:", e) 
            time.sleep(5) 
            continue 
 
        if count == 0: 
            if last_count != 0: 
                os.system('cls' if os.name == 'nt' else 'clear')
                print_big_num("0") 
                last_count = 0 
            time.sleep(5) # wait a bit before checking again
        else: 
            # New mail! Trigger animation
            do_matrix_rain(2) 
            
            # Double check count
            try: 
                count = get_count(f) 
            except: 
                pass 
                
            os.system('cls' if os.name == 'nt' else 'clear')
            print_big_num(str(count)) 
            last_count = count 
            time.sleep(5) 
 
if __name__ == "__main__": 
    run()
