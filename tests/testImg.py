from time import sleep

import cv2
import win32gui
from PIL import ImageGrab

# import pyscreenshot as ImageGrab

coordinate = []


def get_coord(hwnd, ctx):
    if win32gui.IsWindowVisible(hwnd):
        if win32gui.GetClassName(hwnd) == 'OpusApp':
            # win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
            # win32gui.SetForegroundWindow(hwnd)
            sleep(0.5)
            coordinate.clear()
            coordinate.append(win32gui.GetWindowRect(hwnd))


def grab(coords):
    # filename_for_screen = filename.replace(f'.{filename.split(".")[-1]}', '')
    im = ImageGrab.grab(bbox=(coords[0], coords[1] + 150, coords[2], coords[3]))
    im.save('C:\\Python\\CompareIMG\\tests\\test' + '.png', 'PNG')
    # grab(filename + str(number_of_pages), path2


img = cv2.imread('chinese_1.png')
after = cv2.imread('chinese_2.png')
# before_gray = cv2.cvtColor(before, cv2.COLOR_BGR2GRAY)
# after_gray = cv2.cvtColor(after, cv2.COLOR_BGR2GRAY)
#
# cv2.putText(img, 'наш произвольный текст', (50, 300), cv2.FONT_HERSHEY_COMPLEX, 1, color=(0, 255, 0), thickness=2)
# cv2.imshow('Result', img)
# cv2.waitKey()

win32gui.EnumWindows(get_coord, coordinate)
coordinate1 = coordinate[0]
print(coordinate1)

grab(coordinate1)
