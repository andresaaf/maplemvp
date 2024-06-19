from PIL import Image, ImageGrab
import pytesseract
import win32gui
import re
import socket
from time import sleep
from ctypes import windll
import cv2
import numpy as np

def screenshot(x, y, w, h):
    return ImageGrab.grab(bbox=(x, y, w, h))

def get_maplestory_window():
    hwnd = win32gui.FindWindowEx(None, None, None, "MapleStory")
    if hwnd == 0:
        print("NOT FOUND")
        return None
    return win32gui.GetWindowRect(hwnd)

def double_space(img):
    h,w = img.shape[:2]
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]

    pixels = np.sum(thresh, axis=1).tolist()
    space = np.ones((2, w), dtype=np.uint8) * 255
    result = np.zeros((1, w), dtype=np.uint8)

    for i, value in enumerate(pixels):
        if value == 0:
            result = np.concatenate((result, space), axis=0)
        row = gray[i:i+1, 0:w]
        result = np.concatenate((result, row), axis=0)
    return result

def parse_image(img):
    cvimg = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
    cvimg = cv2.resize(cvimg, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
    cvimg = double_space(cvimg)
    #cvimg = cv2.threshold(cvimg, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
    cv2.imwrite("correct.png", cvimg)
    return pytesseract.image_to_string(cvimg)

def parse_mega(text):
    first = True
    mega = []
    for line in text.splitlines():
        if line == '':
            continue
        if True:
            # Next part of chat?
            if line.startswith('Al') and 'Pa' in line and 'Gu' in line:
                break
            first_of_msg = re.match(r'\[(\d{2}):(\d{2})\]', line)
            if first:
                first = False
                # Skip first message if there is no timestamp
                if not first_of_msg:
                    continue
            # If multi-line, merge with previous
            if not first_of_msg and len(mega) > 0:
                mega[-1] += f' {line}'
            else:
                mega.append(line)
    return mega

def parse_mvp(lines):
    mvps = []
    for line in lines:
        l = line.upper()
        time_line = l
        if re.match(r'^\[(\d{2}):(\d{2})\]', l):
            time_line = l[8:]
        time = re.search(r'XX[: ](\d{2})', time_line)
        mvp = re.search(r'MVP', l)
        channel = re.search(r'C[CH]? *(\d{1,2})', l)
        if time and mvp and channel:
            mvps.append({'message_time': line[1:].split(']')[0], 'time': time.group(1), 'channel': channel.group(1), 'message': line})
    return mvps

def filter_mvps(mvps, db):
    new = []
    for mvp in mvps:
        if mvp['message_time'] in db:
            continue
        new.append(mvp)
        db.append(mvp['message_time'])
    return new

def announce(mvps):
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    for mvp in mvps:
        try:
            sock.connect(('192.168.1.2', 8089))
            sock.send(f"{mvp['message_time']}|{mvp['time']}|{mvp['channel']}|{mvp['message']}".encode())
            sock.recv(1)
            sock.close()
        except:
            pass

if __name__ == '__main__':
    # Fix for DPI scaling
    user32 = windll.user32
    user32.SetProcessDPIAware()
    
    db = []
    while True:
        try:
            rect = get_maplestory_window()
            height = rect[3] - rect[1]
            img = screenshot(rect[0], rect[1] + height / 4 + 20, rect[2], rect[3] - height/2 - 30)
            #img = Image.open('Untitled.png')
            text = parse_image(img)
            mega = parse_mega(text)
            mvps = parse_mvp(mega)
            if len(mvps) > 0:
                new_mvps = filter_mvps(mvps, db)
                if len(new_mvps) > 0:
                    announce(new_mvps)
            else:
                db = []
        except Exception as ex:
            print(ex)
            continue
        sleep(2)