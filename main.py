from PIL import Image, ImageGrab
import pytesseract
import win32gui
import re
import socket
from time import sleep

def screenshot(x, y, w, h):
    return ImageGrab.grab(bbox=(x, y, w, h), all_screens=False)

def get_maplestory_window():
    hwnd = win32gui.FindWindowEx(None, None, None, "MapleStory")
    if hwnd == 0:
        print("NOT FOUND")
        return None
    return win32gui.GetWindowRect(hwnd)

def parse_image(img):
    return pytesseract.image_to_string(img)

def parse_mega(text):
    mega_started = False
    first = True
    mega = []
    for line in text.splitlines():
        if line == '':
            continue
        if line == 'Mega' or line == 'Meua':
            mega_started = True
            continue
        if mega_started:
            # Next part of chat?
            if line.startswith('Al') and 'Friend' in line and 'Guild' in line:
                break
            first_of_msg = re.match(r'\[(\d{2}):(\d{2})\]', line)
            if first:
                first = False
                # Skip first message if there is no timestamp
                if not first_of_msg:
                    continue
            # If multi-line, merge with previous
            if not first_of_msg:
                mega[-1] += f' {line}'
            else:
                mega.append(line)
    return mega

def parse_mvp(lines):
    mvps = []
    for line in lines:
        l = line.upper()
        time = re.search(r'XX[:\- ]*(\d{2})', l)
        mvp = re.search(r'MVP', l)
        channel = re.search(r'C[CH] *(\d{1,2})', l)
        if time and mvp and channel:
            mvps.append({'message_time': line[1:].split(']')[0], 'time': time.group(1), 'channel': channel.group(1)})
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
            sock.send(f"{mvp['message_time']}|{mvp['time']}|{mvp['channel']}".encode())
            sock.recv(1)
            sock.close()
        except:
            pass

if __name__ == '__main__':
    db = []
    while True:
        rect = get_maplestory_window()
        img = screenshot(rect[0], rect[1], rect[2], rect[3])
        #img = Image.open('Untitled.png')
        text = parse_image(img)
        mega = parse_mega(text)
        mvps = parse_mvp(mega)
        if len(mvps) > 0:
            new_mvps = filter_mvps(mvps, db)
            if len(new_mvps) > 0:
                print(new_mvps)
                announce(new_mvps)
        else:
            db = []
        sleep(1)