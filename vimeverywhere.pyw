# normal:
# \d*[hjklwW]
# [fFtT].
# ;
# [^$]
# [iIaA]
# r.
# \d*[xs]
# \d*((dd)|(yy)|(cc))
# [DC]
# [oO]
# [dyc]\d*[wW]
# [dyc][^$]
# [dyc][fFtT].
# \.
# :w
# u
# v

# visual:
# \d*[hjkl]
# [^$]
# [xdycs]

import re
import time
import PyHook3
import win32api
import win32con
import pythoncom
import pyperclip
import threading
import PySimpleGUIQt

normal_legal_suffix_re = [re.compile(i) for i in [r'\d+$', r'[fFtTr]$', r'\d*[dyc]$', r'[dyc]\d*$', r'[dyc][fFtT]$', ':']]
normal_functional_re = [re.compile(i) for i in [
    r'\d*([wWhjkl])$', r'[fFtT].$', ';', r'[\^\$]$',
    r'[iIaA]$', r'r.$', r'\d*[xs]$', r'\d*((dd)|(yy)|(cc))$',
    r'[DC]$', r'[oO]$', r'[dyc]\d*[wW]$',
    r'[dyc][\^\$]$', r'[dyc][fFtT].$', r'\.$', ':w', 'u', 'v']]
visual_legal_suffix_re = [re.compile(i) for i in [r'\d+$']]
visual_functional_re = [re.compile(i) for i in [r'\d*[hjkl]$', r'[xdycs]$', r'[\^\$]$']]
state = 'insert'
pause = False
last_ft = ''  # 存储上一次[fFtT].的结果，;命令可以用
last_opt = ''  # 存储上一次删除或复制命令的结果，.命令可以用
command = ''
delay = 0.05
menu = ['BLANK', ['!By Jeremy(Ziming Yuan)', '---', 'Pause', 'Exit']]
tray = PySimpleGUIQt.SystemTray(menu=menu, filename='insert.png', tooltip='VimEverywhere(Insert)')


def to_insert_mode():
    global state
    state = 'insert'
    tray.update(filename='insert.png', tooltip='VimEverywhere(Insert)')


def to_normal_mode():
    global state
    state = 'normal'
    tray.update(filename='normal.png', tooltip='VimEverywhere(Normal)')


def to_visual_mode():
    global state
    state = 'visual'
    tray.update(filename='visual.png', tooltip='VimEverywhere(Visual)')


def pause_and_resume():
    global pause
    pause = not pause
    menu = ['BLANK', ['!By Jeremy', '---', 'Pause √' if pause else 'Pause', 'Exit']]
    tray.update(menu)


def shift_to_home():  # Shift+Home
    win32api.keybd_event(0x10, 0, 0, 0)
    win32api.keybd_event(0x24, 0, win32con.KEYEVENTF_EXTENDEDKEY, 0)
    win32api.keybd_event(0x24, 0, win32con.KEYEVENTF_EXTENDEDKEY | win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(0x10, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(delay)


def shift_to_end():  # Shift+End
    win32api.keybd_event(0x10, 0, 0, 0)
    win32api.keybd_event(0x23, 0, win32con.KEYEVENTF_EXTENDEDKEY, 0)
    win32api.keybd_event(0x23, 0, win32con.KEYEVENTF_EXTENDEDKEY | win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(0x10, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(delay)


def shift_left(num):  # Shift+多次Left
    win32api.keybd_event(0x10, 0, 0, 0)
    left(num)
    win32api.keybd_event(0x10, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(delay)


def shift_right(num):  # Shift+多次Right
    win32api.keybd_event(0x10, 0, 0, 0)
    right(num)
    win32api.keybd_event(0x10, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(delay)


def shift_up(num):  # Shift+多次Up
    win32api.keybd_event(0x10, 0, 0, 0)
    up(num)
    win32api.keybd_event(0x10, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(delay)


def shift_down(num):  # Shift+多次Down
    win32api.keybd_event(0x10, 0, 0, 0)
    down(num)
    win32api.keybd_event(0x10, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(delay)


def control_key(ch):  # Ctrl+字母键
    win32api.keybd_event(0x11, 0, 0, 0)
    win32api.keybd_event(ord(ch.upper()), 0, 0, 0)
    win32api.keybd_event(ord(ch.upper()), 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(0x11, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(delay)


def home():
    win32api.keybd_event(0x24, 0, win32con.KEYEVENTF_EXTENDEDKEY, 0)
    win32api.keybd_event(0x24, 0, win32con.KEYEVENTF_EXTENDEDKEY | win32con.KEYEVENTF_KEYUP, 0)


def end():
    win32api.keybd_event(0x23, 0, win32con.KEYEVENTF_EXTENDEDKEY, 0)
    win32api.keybd_event(0x23, 0, win32con.KEYEVENTF_EXTENDEDKEY | win32con.KEYEVENTF_KEYUP, 0)


def left(num):
    for i in range(num):
        win32api.keybd_event(0x25, 0, win32con.KEYEVENTF_EXTENDEDKEY, 0)
        win32api.keybd_event(0x25, 0, win32con.KEYEVENTF_EXTENDEDKEY | win32con.KEYEVENTF_KEYUP, 0)


def right(num):
    for i in range(num):
        win32api.keybd_event(0x27, 0, win32con.KEYEVENTF_EXTENDEDKEY, 0)
        win32api.keybd_event(0x27, 0, win32con.KEYEVENTF_EXTENDEDKEY | win32con.KEYEVENTF_KEYUP, 0)


def up(num):
    for i in range(num):
        win32api.keybd_event(0x26, 0, win32con.KEYEVENTF_EXTENDEDKEY, 0)
        win32api.keybd_event(0x26, 0, win32con.KEYEVENTF_EXTENDEDKEY | win32con.KEYEVENTF_KEYUP, 0)


def down(num):
    for i in range(num):
        win32api.keybd_event(0x28, 0, win32con.KEYEVENTF_EXTENDEDKEY, 0)
        win32api.keybd_event(0x28, 0, win32con.KEYEVENTF_EXTENDEDKEY | win32con.KEYEVENTF_KEYUP, 0)


def backspace():
    win32api.keybd_event(0x8, 0, 0, 0)
    win32api.keybd_event(0x8, 0, win32con.KEYEVENTF_KEYUP, 0)


def enter():
    win32api.keybd_event(0xD, 0, 0, 0)
    win32api.keybd_event(0xD, 0, win32con.KEYEVENTF_KEYUP, 0)


def get_right_text():  # 获取当前字符到行末的所有字符，注意Normal模式下光标右侧字符为“当前字符”
    last = pyperclip.paste()
    pyperclip.copy('')
    shift_to_end()
    control_key('c')
    now = pyperclip.paste()
    pyperclip.copy(last)
    left(1)
    return now


def get_left_text():  # 获取当前字符到行首的所有字符
    right(1)
    last = pyperclip.paste()
    shift_to_home()
    control_key('c')
    now = pyperclip.paste()
    pyperclip.copy(last)
    right(1)
    left(1)
    return now


def find_lower_word(s):  # 找跳到下一个单词首（单词由同种类型字符组成，并由空格隔开）
    s = str(s)
    if s == '':
        return 0
    else:
        def kind(c):
            return c.isalnum() or c == '_'
        f = s[0]
        for j in range(1, len(s)):
            if s[j].isspace():
                for k in range(j + 1, len(s)):
                    if not s[k].isspace():
                        return k
                return len(s)
            elif kind(f) != kind(s[j]):
                return j
        return len(s)


def find_upper_word(s):  # 找跳到下一个单词首（单词由非空字符组成，并由空格隔开）
    s = str(s)
    if s == '':
        return 0
    else:
        for j in range(1, len(s)):
            if s[j].isspace():
                for k in range(j + 1, len(s)):
                    if not s[k].isspace():
                        return k
                return len(s)
        return len(s)


def find_char(s, c, off):  # 从off开始查找字符c
    if len(s) <= off:
        return 0
    for i, j in enumerate(s[off:]):
        if j == c:
            return i + off
    return 0


def process_command(cmd):  # 处理命令
    global normal_functional_re, state, last_ft, last_opt, command
    if normal_functional_re[0].match(cmd):  # \d*[hjklwW]
        nums, opt = cmd[:-1], cmd[-1]
        if nums == '':
            num = 1
        else:
            num = int(nums)
        if opt == 'w' or opt == 'W':
            s = get_right_text()
            cnt = 0
            for i in range(num):
                offset = find_lower_word(s) if opt == 'w' else find_upper_word(s)
                s = s[offset:]
                cnt += offset
            print(cnt)
            right(cnt)
        if opt == 'h':
            left(num)
        if opt == 'j':
            down(num)
        if opt == 'k':
            up(num)
        if opt == 'l':
            right(num)
    if normal_functional_re[1].match(cmd):  # [fFtT].
        opt, ch = tuple(list(cmd))
        if opt == 'f' or opt == 't':
            s = get_right_text()
            right(find_char(s, ch, 1) if opt == 'f' else max(find_char(s, ch, 2) - 1, 0))
        if opt == 'F' or opt == 'T':
            s = get_left_text()[::-1]
            left(find_char(s, ch, 1) if opt == 'F' else max(find_char(s, ch, 2) - 1, 0))
        last_ft = cmd
    if normal_functional_re[2].match(cmd):  # ;
        process_command(last_ft)
    if normal_functional_re[3].match(cmd):  # [\^\$]
        if cmd == '^':
            win32api.keybd_event(0x10, 0, win32con.KEYEVENTF_KEYUP, 0)  # 这个键按下的时候需要按Shift，所以要先将Shift搞掉
            home()
        else:
            win32api.keybd_event(0x10, 0, win32con.KEYEVENTF_KEYUP, 0)
            end()
            left(1)
    if normal_functional_re[4].match(cmd):  # [aAiI]
        if cmd == 'a':
            right(1)
        if cmd == 'I':
            win32api.keybd_event(0x10, 0, win32con.KEYEVENTF_KEYUP, 0)
            home()
        if cmd == 'A':
            win32api.keybd_event(0x10, 0, win32con.KEYEVENTF_KEYUP, 0)
            end()
        to_insert_mode()
    if normal_functional_re[5].match(cmd):  # r.
        ch = cmd[1]
        right(1)
        backspace()
        last = pyperclip.paste()
        pyperclip.copy(ch)
        control_key('v')
        pyperclip.copy(last)
        last_opt = cmd
    if normal_functional_re[6].match(cmd):  # \d*[xs]
        nums, opt = cmd[:-1], cmd[-1]
        if nums == '':
            num = 1
        else:
            num = int(nums)
        num = min(num, len(get_right_text()))
        shift_right(num)
        control_key('x')
        if opt == 's':
            to_insert_mode()
        last_opt = cmd
    if normal_functional_re[7].match(cmd):  # \d*((dd)|(yy)|(cc))
        nums, opt = cmd[:-2], cmd[-2:]
        if nums == '':
            num = 1
        else:
            num = int(nums)
        if opt == 'dd' or opt == 'cc':
            paste = ''
            for i in range(num):
                home()
                shift_to_end()
                control_key('x')
                paste += pyperclip.paste() + '\n'
                if len(get_right_text()) == 0:
                    down(1)
                    backspace()
            home()
            pyperclip.copy(paste)
            if opt == 'cc':
                to_insert_mode()
            last_opt = cmd
        else:
            p = len(get_left_text()) - 1
            paste = ''
            home()
            for i in range(num):
                shift_to_end()
                control_key('c')
                paste += pyperclip.paste() + '\n'
                left(1)
                down(1)
            up(num)
            right(p)
            pyperclip.copy(paste)
    if normal_functional_re[8].match(cmd):  # [DC]
        shift_to_end()
        control_key('x')
        if cmd == 'C':
            to_insert_mode()
        last_opt = cmd
    if normal_functional_re[9].match(cmd):  # [oO]
        if cmd == 'o':
            end()
            enter()
            to_insert_mode()
        else:
            win32api.keybd_event(0x10, 0, win32con.KEYEVENTF_KEYUP, 0)
            home()
            enter()
            up(1)
            to_insert_mode()
        last_opt = cmd
    if normal_functional_re[10].match(cmd):  # [dyc]\d*[wW]
        opt, nums, pos = cmd[0], cmd[1:-1], cmd[-1]
        if nums == '':
            num = 1
        else:
            num = int(nums)
        s = get_right_text()
        o = s
        cnt = 0
        for i in range(num):
            offset = find_lower_word(s) if pos == 'w' else find_upper_word(s)
            s = s[offset:]
            cnt += offset
        if opt == 'c' or opt == 'y':  # c和y不能把空格也弄进去
            while cnt > 0 and o[cnt - 1].isspace():
                cnt -= 1
        shift_right(cnt)
        if opt == 'd':
            control_key('x')
            last_opt = cmd
        if opt == 'c':
            control_key('x')
            to_insert_mode()
            last_opt = cmd
        if opt == 'y':
            control_key('c')
            left(1)
    if normal_functional_re[11].match(cmd):  # [dyc][\^\$]
        opt, pos = cmd[0], cmd[1]
        shift_to_home() if pos == '^' else shift_to_end()
        if opt == 'd':
            control_key('x')
            last_opt = cmd
        if opt == 'c':
            control_key('x')
            to_insert_mode()
            last_opt = cmd
        if opt == 'y':
            control_key('c')
            right(1) if pos == '^' else left(1)
    if normal_functional_re[12].match(cmd):  # [dyc][fFtT].
        opt, pos, ch = tuple(list(cmd))
        if pos == 'f' or pos == 't':
            s = get_right_text()
            shift_right((find_char(s, ch, 1) + 1) if pos == 'f' else find_char(s, ch, 2))
        if pos == 'F' or pos == 'T':
            s = get_left_text()[::-1]
            shift_left(find_char(s, ch, 1) if pos == 'F' else max(find_char(s, ch, 2) - 1, 0))
        if opt == 'd':
            control_key('x')
            last_opt = cmd
        if opt == 'c':
            control_key('x')
            to_insert_mode()
            last_opt = cmd
        if opt == 'y':
            control_key('c')
            if pos == 'f' or pos == 't':
                left(1)
            if pos == 'F' or pos == 'T':
                right(1)
        last_ft = cmd[1:]
    if normal_functional_re[13].match(cmd):  # \.
        process_command(last_opt)
    if normal_functional_re[14].match(cmd):  # :w
        control_key('s')
    if normal_functional_re[15].match(cmd):  # u
        control_key('z')
    if normal_functional_re[16].match(cmd):  # v
        to_visual_mode()


def process_vcommand(cmd):  # 处理命令
    global visual_functional_re, state, command
    if visual_functional_re[0].match(cmd):  # \d*[hjkl]
        nums, opt = cmd[:-1], cmd[-1]
        if nums == '':
            num = 1
        else:
            num = int(nums)
        if opt == 'h':
            shift_left(num)
        if opt == 'j':
            shift_down(num)
        if opt == 'k':
            shift_up(num)
        if opt == 'l':
            shift_right(num)
    if visual_functional_re[1].match(cmd):  # [xdycs]
        if cmd == 'x' or cmd == 'd':
            control_key('x')
            to_normal_mode()
        if cmd == 'y':
            control_key('c')
            to_normal_mode()
        if cmd == 'c' or cmd == 's':
            control_key('x')
            to_insert_mode()
    if visual_functional_re[2].match(cmd):  # [^$]
        if cmd == '^':
            shift_to_home()
        if cmd == '$':
            shift_to_end()


def process_normal(c):
    global normal_legal_suffix_re, normal_functional_re, command
    if c == 0x1B:  # 按Esc清空命令字符串
        command = ''
        return False
    if c < 0x20:  # 无视控制键
        return True
    else:
        command += chr(c)  # 命令字符串加入当前字符
        print(command)
        if all(map(lambda x: x.match(command) is None, normal_legal_suffix_re + normal_functional_re)):
            command = ''  # 命令字符串不是合法命令及其前缀
            return False
        elif any(map(lambda x: x.match(command) is not None, normal_functional_re)):  # 合法命令
            threading.Thread(target=process_command, args=(command,)).start()
            # 处理命令的过程中有按键后的delay，所以必须异步处理，否则太晚return False命令会被编辑器捕获到
            command = ''
            return False
    return False


def process_insert(c):
    global state
    if c == 0x1B:  # 按下Esc键返回正常模式
        to_normal_mode()
        return False
    return True


def process_visual(c):
    global visual_legal_suffix_re, visual_functional_re, command
    if c == 0x1B:  # 按Esc回到正常模式
        command = ''
        to_normal_mode()
        return False
    if c < 0x20:  # 无视控制键
        return True
    else:
        command += chr(c)  # 命令字符串加入当前字符
        print(command)
        if all(map(lambda x: x.match(command) is None, visual_legal_suffix_re + visual_functional_re)):
            command = ''  # 命令字符串不是合法命令及其前缀
            return False
        elif any(map(lambda x: x.match(command) is not None, visual_functional_re)):  # 合法命令
            threading.Thread(target=process_vcommand, args=(command,)).start()
            # 处理命令的过程中有按键后的delay，所以必须异步处理，否则太晚return False命令会被编辑器捕获到
            command = ''
            return False
    return False


def on_keyboard_event(event):
    if not pause:
        if state == 'normal':
            return process_normal(event.Ascii)
        if state == 'insert':
            return process_insert(event.Ascii)
        if state == 'visual':
            return process_visual(event.Ascii)
    return True


hm = PyHook3.HookManager()
hm.KeyDown = on_keyboard_event
hm.HookKeyboard()
while True:
    pythoncom.PumpWaitingMessages()  # 这个函数不会阻塞
    item = tray.Read(0)
    if item == 'Pause' or item == 'Pause √':
        pause_and_resume()
    if item == 'Exit':
        tray.hide()
        break
