import os
import shlex
import subprocess as sb
from time import sleep

import cv2
import numpy as np
import pyautogui as pg
from PIL import ImageGrab
from skimage.metrics import structural_similarity

list_of_files = os.getcwd() + '/data/files/'
ppt_path = os.getcwd() + '/data/before/'
pptx_path = os.getcwd() + '/data/after/'
ppt_path_imgmagic = 'data/before/'
pptx_path_imgmagic = 'data/after/'


def compare_img(pathppt, pathpptx, file_name):
    before = cv2.imread(pathppt)
    after = cv2.imread(pathpptx)
    # Convert images to grayscale
    before_gray = cv2.cvtColor(before, cv2.COLOR_BGR2GRAY)
    after_gray = cv2.cvtColor(after, cv2.COLOR_BGR2GRAY)
    # Compute SSIM between two images
    (score, diff) = structural_similarity(before_gray, after_gray, full=True)
    print("Image similarity", score * 100)

    # The diff image contains the actual image differences between the two images
    # and is represented as a floating point data type in the range [0,1]
    # so we must convert the array to 8-bit unsigned integers in the range
    # [0,255] before we can use it with OpenCV
    diff = (diff * 255).astype("uint8")

    # Threshold the difference image, followed by finding contours to
    # obtain the regions of the two input images that differ
    thresh = cv2.threshold(diff, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)[1]
    contours = cv2.findContours(thresh.copy(), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    contours = contours[0] if len(contours) == 2 else contours[1]

    mask = np.zeros(before.shape, dtype='uint8')
    filled_after = after.copy()

    for c in contours:
        area = cv2.contourArea(c)
        if area > 40:
            x, y, w, h = cv2.boundingRect(c)
            cv2.rectangle(before, (x, y), (x + w, y + h), (0, 0, 255), 0)
            cv2.rectangle(after, (x, y), (x + w, y + h), (0, 0, 255), 0)
            cv2.drawContours(mask, [c], 0, (0, 0, 255), 0)
            cv2.drawContours(filled_after, [c], 0, (0, 0, 255), 0)
    try:
        folder = file_name.split('.')
        os.mkdir(os.getcwd() + f'\\data/{folder[0]}')
    except Exception:
        print(Exception, 'folder created')

    cv2.imwrite(os.getcwd() + f'\\data/{folder[0]}/{file_name}_before_ppt.png', before)
    cv2.imwrite(os.getcwd() + f'\\data/{folder[0]}/{file_name}_after_pptx.png', after)
    # cv2.imwrite(os.getcwd() + f'\\data/{number}diff.png', diff)
    # cv2.imwrite(os.getcwd() + f'\\data/{number}_mask.png', mask)
    cv2.imwrite(os.getcwd() + f'\\data/{folder[0]}/{file_name}_filled_after.png', filled_after)

    # cv2.imshow('before', before)
    # cv2.imshow('before', after)
    # cv2.imshow('diff', diff)
    # cv2.imshow('mask', mask)
    # cv2.imshow('filled before', filled_after)
    # cv2.waitKey(0)


def compare_img_wirh_imagemagic(pathppt, pathpptx, file_name):
    try:
        folder = file_name.split('.')
        folder = folder[0] + '_imgmagic'
        os.mkdir(os.getcwd() + f'\\data/{folder}')

    except Exception:
        print(Exception, 'folder created')
    command = f'magick compare -compose src {pathppt} {pathpptx} data/{folder}/{file_name}.png'
    sb.call(shlex.split(command), shell=True)
    pass


# def grab(filename, path):
#     im = ImageGrab.grab()
#     im.save(path + filename + '.png', 'PNG')
#     return im


def run(file_name):
    sb.Popen([r"C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE", f"./data/files/{file_name}"])
    sleep(3)


def grab(filename, path):
    number = 0
    for el in range(5):
        im = ImageGrab.grab()
        im.save(path + filename + str(number) + '.png', 'PNG')
        # grab(filename + str(number), path2)
        pg.press('down')
        sleep(0.3)
        number += 1


def go():
    list_f = os.listdir(list_of_files)
    for file_name in list_f:
        # index = file_name.find('after')
        index = file_name.split('.')
        print(index)
        if index[1] == 'after':
            run(file_name)
            sleep(3)
            file_name = file_name.replace('after', '')
            grab(file_name, pptx_path)
            sb.call(["taskkill", "/IM", "POWERPNT.EXE"])
        elif index[1] == 'before':
            run(file_name)
            sleep(3)
            file_name = file_name.replace('before', '')
            grab(file_name, ppt_path)
            sb.call(["taskkill", "/IM", "POWERPNT.EXE"])


if __name__ == "__main__":
    # go()
    mass = os.listdir(ppt_path)
    for filename in mass:
        compare_img(ppt_path + filename, pptx_path + filename, filename)
        # os.remove(ppt_path + filename)
        # os.remove(pptx_path + filename)
        # print(ppt_path + filename)
        # compare_img_wirh_imagemagic(ppt_path_imgmagic + filename, pptx_path_imgmagic + filename, filename)
       