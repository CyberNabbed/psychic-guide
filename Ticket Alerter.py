#!/usr/bin/env python 
# -*- coding: utf-8 -*- 
 
""" 
CLI monitor for unread e-mails in a specific Outlook folder. 
Press **Enter** at any time and every remaining unread mail in that folder 
is marked read immediately. 
 
Dependencies (Windows only): 
    pip install pywin32 
    pip install colorama    # optional, for green Matrix text 
""" 
 
import time 
import random 
import os 
import sys 
 
# ====================================================================
# USER CONFIGURATION
# ====================================================================
# REPLACE THE TEXT BELOW with the specific folder name you want to monitor
TARGET_FOLDER_NAME = "YOUR_TARGET_FOLDER_NAME_HERE" 
# Example: "Helpdesk Tickets", "Invoices", "Alerts"
# ====================================================================

# -------------------------------------------------------------------- 
# Outlook COM support 
# -------------------------------------------------------------------- 
try: 
    import win32com.client 
except ImportError: 
    print("Error: pywin32 is not installed. Install it with `pip install pywin32`.") 
    sys.exit(1) 
 
# -------------------------------------------------------------------- 
# Windows-only, non-blocking keyboard input 
# -------------------------------------------------------------------- 
try: 
    import msvcrt          # built-in on Windows 
except ImportError: 
    print("This script must be run on Windows (msvcrt not found).") 
    sys.exit(1) 
 
# -------------------------------------------------------------------- 
# Optional colour output 
# -------------------------------------------------------------------- 
try: 
    from colorama import init, Fore, Style 
    init() 
    USE_COLORAMA = True 
except ImportError: 
    USE_COLORAMA = False 
 
# -------------------------------------------------------------------- 
# Matrix animation width 
# -------------------------------------------------------------------- 
MATRIX_WIDTH = len("44444444444444444444444444444444444444444444") 
 
# -------------------------------------------------------------------- 
# Large ASCII digits (20 rows each) 
# -------------------------------------------------------------------- 
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
 
# -------------------------------------------------------------------- 
# Helper to widen a line (visual scaling) 
# -------------------------------------------------------------------- 
def widen_line(line, factor=2): 
    return "".join(c * factor for c in line) 
 
# -------------------------------------------------------------------- 
# Display big digits horizontally 
# -------------------------------------------------------------------- 
def display_large_number(num_str, widen_factor=2): 
    art_per_digit = [ 
        [widen_line(row, widen_factor) for row in BIG_DIGITS[d]] 
        for d in num_str 
    ] 
    for row in range(20): 
        print("  ".join(art[row] for art in art_per_digit)) 
 
# -------------------------------------------------------------------- 
# Matrix rain effect 
# -------------------------------------------------------------------- 
def matrix_animation(duration=2.0): 
    start = time.time() 
    while time.time() - start < duration: 
        line = "" 
        for _ in range(MATRIX_WIDTH): 
            ch = random.choice("01X|/\\{}[]()#$%^&*+;:ABCDEFGHIJKLMNOPQRSTUVWXYZ") 
            if USE_COLORAMA: 
                line += Fore.GREEN + ch + Style.RESET_ALL 
            else: 
                line += ch 
        print(line) 
        time.sleep(0.05) 
    clear_screen() 
 
# -------------------------------------------------------------------- 
# Clear screen 
# -------------------------------------------------------------------- 
def clear_screen(): 
    os.system('cls' if os.name == 'nt' else 'clear') 
 
# -------------------------------------------------------------------- 
# Find the folder across all mailbox roots 
# -------------------------------------------------------------------- 
def find_target_folder(folder_name=TARGET_FOLDER_NAME): 
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI") 
    for i in range(1, outlook.Folders.Count + 1): 
        root = outlook.Folders.Item(i) 
        try: 
            return root.Folders.Item(folder_name) 
        except Exception: 
            pass 
    raise RuntimeError(f"Folder “{folder_name}” not found under any mailbox root.") 
 
# -------------------------------------------------------------------- 
# Get unread count (forces refresh via Restrict) 
# -------------------------------------------------------------------- 
def get_unread_count(folder): 
    try: 
        return folder.Items.Restrict("[UnRead] = True").Count 
    except Exception: 
        return folder.UnreadItemCount 
 
# -------------------------------------------------------------------- 
# Mark all unread mails as read – copy items first to avoid collection mutation 
# -------------------------------------------------------------------- 
def mark_all_unread_as_read(folder): 
    try: 
        unread_items = folder.Items.Restrict("[UnRead] = True") 
        # Copy to a list so the collection isn't altered while iterating 
        to_mark = [item for item in unread_items] 
        for item in to_mark: 
            if item.UnRead: 
                item.UnRead = False 
                # No explicit Save() needed; setting UnRead saves automatically 
    except Exception as exc: 
        # Most COM errors are transient; surface the message but keep running 
        clear_screen() 
        print("Error while marking as read:", exc) 
        time.sleep(2) 
 
# -------------------------------------------------------------------- 
# Main loop 
# -------------------------------------------------------------------- 
def main_loop(): 
    last_displayed = None 
 
    while True: 
        # ------------------------------------------------------------ 
        # 1. Keyboard check – press Enter to mark all as read 
        # ------------------------------------------------------------ 
        if msvcrt.kbhit(): 
            key = msvcrt.getwch() 
            if key == '\r':                                   # Enter 
                try: 
                    folder = find_target_folder() 
                    mark_all_unread_as_read(folder) 
                    clear_screen() 
                    display_large_number("0") 
                    last_displayed = 0 
                except Exception as exc: 
                    clear_screen() 
                    print("Error:", exc) 
                time.sleep(2) 
                continue   # restart the loop with fresh counts 
 
        # ------------------------------------------------------------ 
        # 2. Display logic 
        # ------------------------------------------------------------ 
        try: 
            folder = find_target_folder() 
            unread = get_unread_count(folder) 
        except Exception as exc: 
            clear_screen() 
            print("Error accessing Outlook folder:", exc) 
            time.sleep(10) 
            continue 
 
        if unread == 0: 
            if last_displayed != 0: 
                clear_screen() 
                display_large_number("0") 
                last_displayed = 0 
            time.sleep(10) 
        else: 
            matrix_animation(2) 
            # Refresh the count right after the animation 
            try: 
                unread = get_unread_count(folder) 
            except Exception: 
                pass 
            clear_screen() 
            display_large_number(str(unread)) 
            last_displayed = unread 
            time.sleep(5) 
 
# -------------------------------------------------------------------- 
# Entry point 
# -------------------------------------------------------------------- 
def run(): 
    clear_screen() 
    main_loop() 
 
if __name__ == "__main__": 
    run()
