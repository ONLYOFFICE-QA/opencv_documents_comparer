import json
import subprocess as sb
from time import sleep

import cv2
import numpy as np
import pyautogui as pg
from PIL import ImageGrab
from rich import print
from rich.progress import track
from skimage.metrics import structural_similarity
from win32com.client import Dispatch

from var import *


def compare_img(path_to_image_befor_conversion, path_to_image_after_conversion, file_name):
    before = cv2.imread(path_to_image_befor_conversion)
    after = cv2.imread(path_to_image_after_conversion)
    # Convert images to grayscale
    before_gray = cv2.cvtColor(before, cv2.COLOR_BGR2GRAY)
    after_gray = cv2.cvtColor(after, cv2.COLOR_BGR2GRAY)
    # Compute SSIM between two images
    (score, diff) = structural_similarity(before_gray, after_gray, full=True)

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

    folder = file_name.split('_')[0]
    if not os.path.exists(f'{path_to_result}{folder}'):
        os.mkdir(f'{path_to_result}{folder}')

    file_name = file_name.split('.')[0]
    sheet = file_name.split('_')[-1]
    file_name_for_print = file_name.split('_')[0]
    print(f"{file_name_for_print} Sheet:{sheet} similarity", score * 100)
    cv2.imwrite(f'{path_to_result}{folder}/{file_name}_before_conv.png', before)
    cv2.imwrite(f'{path_to_result}{folder}/{file_name}_after_conv.png', after)
    # cv2.imwrite(os.getcwd() + f'\\data/{number_of_pages}diff.png', diff)
    # cv2.imwrite(os.getcwd() + f'\\data/{number_of_pages}_mask.png', mask)
    cv2.imwrite(f'{path_to_result}{folder}/{file_name}_filled_after.png', filled_after)

    # cv2.imshow('before', before)
    # cv2.imshow('before', after)
    # cv2.imshow('diff', diff)
    # cv2.imshow('mask', mask)
    # cv2.imshow('filled before', filled_after)
    # cv2.waitKey(0)


def run(file_name):
    sb.Popen([r"C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE", f"./data/files/{file_name}"])


def grab(path, filename, number_of_pages):
    sleep(wait_for_open)
    num_of_sheet = 1
    for el in range(int(number_of_pages)):
        im = ImageGrab.grab()
        im.save(path + filename + str(f'_{num_of_sheet}') + '.png', 'PNG')
        # grab(filename + str(number_of_pages), path2)
        pg.click()
        pg.press('pgdn')
        sleep(0.3)
        num_of_sheet += 1


def create_word_json(word, file_name_for_save, path_for_save):
    statistics_word = {
        'num_of_sheets': f'{word.ComputeStatistics(2)}',
        'number_of_lines': f'{word.ComputeStatistics(1)}',
        'word_count': f'{word.ComputeStatistics(0)}',
        'number_of_characters_without_spaces': f'{word.ComputeStatistics(3)}',
        'number_of_characters_with_spaces': f'{word.ComputeStatistics(5)}',
        'number_of_abzad': f'{word.ComputeStatistics(4)}',
    }
    with open(f'{path_for_save}{file_name_for_save}_doc.json', 'w') as f:
        json.dump(statistics_word, f)
    return statistics_word


def word_opener(file_name, path_for_open, path_for_save):
    word = Dispatch('Word.Application')
    word.Visible = True
    file_name_for_screen = file_name.split('.')[0]
    # print(f'{custom_path_to_document_from}{file_name}')
    word = word.Documents.Open(f'{path_for_open}{file_name}')
    word.Repaginate()
    statistics_word = create_word_json(word, file_name_for_screen, path_for_save)
    return word, statistics_word


def open_document_for_compare(list_of_files, from_extension, to_extension):
    for file_name in list_of_files:
        extension = file_name.split('.')[-1]
        if extension == 'ppt' or extension == 'pptx':
            run(file_name, 'POWERPNT.EXE')

        elif extension == 'doc' or 'docx':
            word, statistics_word = word_opener(file_name, custom_path_to_document_to, path_to_tmpimg_after_conversion)
            grab(path_to_tmpimg_after_conversion, file_name.split('.')[0], statistics_word['num_of_sheets'])
            word.Close()
            file_name_from = file_name.replace(f'.{to_extension}', f'.{from_extension}')
            word, statistics_word = word_opener(file_name_from, custom_path_to_document_from,
                                                path_to_tmpimg_befor_conversion)
            word.Close()
    # sb.call(["taskkill", "/IM", "WINWORD.EXE"])
    # sb.call(["taskkill", "/IM", "WINWORD.EXE"])
    # sb.call(f'powershell.exe kill -Name WINWORD', shell=True)


def create_project_dirs():
    if not os.path.exists(path_to_tmpimg_befor_conversion):
        os.mkdir(path_to_tmpimg_befor_conversion)

    if not os.path.exists(path_to_tmpimg_after_conversion):
        os.mkdir(path_to_tmpimg_after_conversion)

    if not os.path.exists(path_to_result):
        os.mkdir(path_to_result)


def start_to_compare_images():
    list_of_images = os.listdir(path_to_tmpimg_befor_conversion)
    for image_name in track(list_of_images, description='Comparing In Progress'):
        if image_name.split('.')[-1] != 'json':
            if os.path.exists(f'{path_to_tmpimg_befor_conversion}{image_name}') and os.path.exists(
                    f'{path_to_tmpimg_after_conversion}{image_name}'):
                compare_img(path_to_tmpimg_befor_conversion + image_name,
                            path_to_tmpimg_after_conversion + image_name,
                            image_name)
            else:
                print(f'[bold red] FIle not found [/bold red]{image_name}')
        else:
            print(f'[bold red] JSON OR WATH? [/bold red]{image_name}')
    # sb.call(f'powershell.exe rm {path_to_tmpimg_befor_conversion}*,'
    #         f'{path_to_tmpimg_after_conversion}* -Recurse',
    #         shell=True)


def dict_compare(d1, d2):
    d1_keys = set(d1.keys())
    d2_keys = set(d2.keys())
    shared_keys = d1_keys.intersection(d2_keys)
    modified = {o: (f'Befor {d1[o]}', f'After {d2[o]}') for o in shared_keys if d1[o] != d2[o]}
    return modified


def run_compare_statistics():
    a = json.load(open(path_to_result + 'Alimentaire_Etude_Planning_StrategiqueKM_2010_doc.json'))
    b = json.load(open(path_to_result + 'Alimentaire_Etude_Planning_StrategiqueKM_2010_docx.json'))

    modified = dict_compare(a, b)
    print(modified)


if __name__ == "__main__":
    # sb.call(f'powershell.exe rm {path_to_tmpimg_befor_conversion}*,'
    #         f'{path_to_tmpimg_after_conversion}* -Recurse',
    #         shell=True)
    #
    # create_project_dirs()
    # open_document_for_compare(list_file_names_doc_from_compare, extension_from, extension_to)

    start_to_compare_images()
    # sb.call(f'powershell.exe rm {path_to_tmpimg_befor_conversion}*,'
    #         f'{path_to_tmpimg_after_conversion}* -Recurse',
    #         shell=True)

    # list_of_files = os.listdir(path_to_compare_files)
