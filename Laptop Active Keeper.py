"""Keep Laptop and CG Active v3.0 Maze Inside Out"""
from datetime import datetime
from ctypes import *
from time import sleep
import pyautogui as pag
import os


def main():
    now = datetime.now()
    start_time = now.strftime("%H:%M:%S")
    os.system('color 0A')
    os.system('cls')
    print(
        f"** Active Keeper v4.0 **\n\nStarted Time: {start_time}\nCurrent Time: \n\nTime Passed : ")
    x, y = 250, 1050

    old_x, old_y = pag.position()

    while True:
        now = datetime.now()
        current_time = now.strftime("%H:%M:%S")
        print_at(3, 14, current_time)
        time_passed = time_diff_cal(start_time, current_time)
        print_at(5, 14, time_passed)
        if int(time_passed[6:]) == 0 and int(time_passed[3:5]) % 4 == 0:
            current_x, current_y = pag.position()
            if old_x == current_x and old_y == current_y:
                pag.click(x, y)
                pag.moveTo(current_x, current_y)
            old_x, old_y = current_x, current_y
            sleep(1)


def time_diff_cal(start, current):
    start = [int(x) for x in start.split(':')]
    current = [int(x) for x in current.split(':')]
    start_seconds = (start[0] * 60 + start[1]) * 60 + start[2]
    current_seconds = (current[0] * 60 + current[1]) * 60 + current[2]
    diff = current_seconds - start_seconds

    hours = diff//3600
    minutes = (diff - 3600*hours)//60
    seconds = diff - 3600*hours - 60*minutes

    if hours < 10:
        hours = f"0{hours}"
    if minutes < 10:
        minutes = f"0{minutes}"
    if seconds < 10:
        seconds = f"0{seconds}"

    return f"{hours}:{minutes}:{seconds}"


def print_at(r, c, s):
    h = windll.kernel32.GetStdHandle(STD_OUTPUT_HANDLE)
    windll.kernel32.SetConsoleCursorPosition(h, COORD(c, r))
    c = s.encode("windows-1252")
    windll.kernel32.WriteConsoleA(h, c_char_p(c), len(c), None, None)


STD_OUTPUT_HANDLE = -11


class COORD(Structure):
    pass


COORD._fields_ = [("X", c_short), ("Y", c_short)]
os.system('mode 25,8')
main()
