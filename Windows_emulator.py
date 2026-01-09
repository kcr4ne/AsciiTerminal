
import sys
import time
import os
import ctypes
from ctypes import wintypes

# ------------------------------------------------------------------
# Ctypes Struct Definitions for SendInput
# ------------------------------------------------------------------
user32 = ctypes.windll.user32

INPUT_MOUSE    = 0
INPUT_KEYBOARD = 1
INPUT_HARDWARE = 2

KEYEVENTF_EXTENDEDKEY = 0x0001
KEYEVENTF_KEYUP       = 0x0002
KEYEVENTF_UNICODE     = 0x0004
KEYEVENTF_SCANCODE    = 0x0008

VK_RETURN = 0x0D
VK_LWIN   = 0x5B
VK_SHIFT  = 0x10
VK_CONTROL= 0x11

class KEYBDINPUT(ctypes.Structure):
    _fields_ = (("wVk",         wintypes.WORD),
                ("wScan",       wintypes.WORD),
                ("dwFlags",     wintypes.DWORD),
                ("time",        wintypes.DWORD),
                ("dwExtraInfo", ctypes.c_ulonglong))

class MOUSEINPUT(ctypes.Structure):
    _fields_ = (("dx",          wintypes.LONG),
                ("dy",          wintypes.LONG),
                ("mouseData",   wintypes.DWORD),
                ("dwFlags",     wintypes.DWORD),
                ("time",        wintypes.DWORD),
                ("dwExtraInfo", ctypes.c_ulonglong))

class HARDWAREINPUT(ctypes.Structure):
    _fields_ = (("uMsg",    wintypes.DWORD),
                ("wParamL", wintypes.WORD),
                ("wParamH", wintypes.WORD))

class INPUT(ctypes.Structure):
    class _INPUT(ctypes.Union):
        _fields_ = (("ki", KEYBDINPUT),
                    ("mi", MOUSEINPUT),
                    ("hi", HARDWAREINPUT))
    _anonymous_ = ("_input",)
    _fields_ = (("type",   wintypes.DWORD),
                ("_input", _INPUT))

LPINPUT = ctypes.POINTER(INPUT)

def _send_input(inputs):
    nInputs = len(inputs)
    lpInputs = (INPUT * nInputs)(*inputs)
    pInputs = ctypes.cast(lpInputs, LPINPUT)
    user32.SendInput(nInputs, pInputs, ctypes.sizeof(INPUT))

def press_key(vk_code):
    x = INPUT(type=INPUT_KEYBOARD,
              ki=KEYBDINPUT(wVk=vk_code))
    _send_input([x])
    time.sleep(0.02) # 안정성을 위한 딜레이

def release_key(vk_code):
    x = INPUT(type=INPUT_KEYBOARD,
              ki=KEYBDINPUT(wVk=vk_code,
                            dwFlags=KEYEVENTF_KEYUP))
    _send_input([x])
    time.sleep(0.02)

def type_char_unicode(c):
    # Send unicode character
    x = INPUT(type=INPUT_KEYBOARD,
              ki=KEYBDINPUT(wVk=0,
                            wScan=ord(c),
                            dwFlags=KEYEVENTF_UNICODE))
    y = INPUT(type=INPUT_KEYBOARD,
              ki=KEYBDINPUT(wVk=0,
                            wScan=ord(c),
                            dwFlags=KEYEVENTF_UNICODE | KEYEVENTF_KEYUP))
    _send_input([x, y])
    time.sleep(0.01) # 타이핑 속도 조절

# ------------------------------------------------------------------
# DuckyScript Logic
# ------------------------------------------------------------------

def process_line(line):
    line = line.strip()
    if not line or line.startswith("REM"):
        return

    parts = line.split(" ", 1)
    command = parts[0].upper()
    args = parts[1] if len(parts) > 1 else ""

    if command == "DELAY":
        ms = int(args)
        print(f"Waiting {ms}ms...")
        time.sleep(ms / 1000.0)

    elif command == "STRING":
        print(f"Typing: {args}")
        for char in args:
            type_char_unicode(char)

    elif command == "ENTER":
        print("[Cmd] Pressing ENTER")
        press_key(VK_RETURN)
        release_key(VK_RETURN)

    elif command == "GUI":
        # Handle GUI r, GUI d, etc.
        key = args.lower()
        print(f"[Cmd] Pressing Win + {key}")
        
        # Win 키 누름
        press_key(VK_LWIN)
        time.sleep(0.2) # 핫키 인식 대기 (늘림)
        
        if key:
            # 문자에 해당하는 VK Code 찾기 (정확한 매핑)
            # VkKeyScanW returns: low byte = VK code, high byte = shift state
            vk_res = user32.VkKeyScanW(ord(key))
            vk_code = vk_res & 0xFF
            
            if vk_code != -1:
                press_key(vk_code)
                time.sleep(0.1)
                release_key(vk_code)
            else:
                # 매핑 실패 시 fallback
                type_char_unicode(key)

        # Win 키 뗌
        time.sleep(0.2)
        release_key(VK_LWIN)
    
    else:
        print(f"Warning: Unknown command {command}")

def run_ducky_script(filepath):
    if not os.path.exists(filepath):
        print(f"File not found: {filepath}")
        return

    print(f"Running DuckyScript: {filepath}")
    print("WARNING: This will simulate keyboard input. Hands off!")
    time.sleep(2) # Give user a moment

    with open(filepath, 'r', encoding='utf-8') as f:
        for line in f:
            process_line(line)

    print("Done.")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python run_ducky_on_windows.py <duckyscript_file>")
    else:
        run_ducky_script(sys.argv[1])
