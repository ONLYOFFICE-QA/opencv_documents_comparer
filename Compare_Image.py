from time import sleep
import imageio

import cv2
import numpy as np
import pyautogui as pg
from PIL import ImageGrab
from skimage.metrics import structural_similarity

from helper import *
from var import *


class CompareImage:
    def __init__(self):
        self.start_to_compare_images()

    @staticmethod
    def compare_img(path_to_image_befor_conversion, path_to_image_after_conversion, file_name):
        before = cv2.imread(path_to_image_befor_conversion)
        after = cv2.imread(path_to_image_after_conversion)
        before_gray = cv2.cvtColor(before, cv2.COLOR_BGR2GRAY)
        after_gray = cv2.cvtColor(after, cv2.COLOR_BGR2GRAY)

        (score, diff) = structural_similarity(before_gray, after_gray, full=True)

        diff = (diff * 255).astype("uint8")

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

        # folder_name = file_name.split('.')[0]
        folder_name = file_name.replace(f'.{file_name.split(".")[-1]}', '')
        folder_name = file_name.replace(f'_{file_name.split("_")[-1]}', '')
        Helper.create_dir(f'{path_to_result}{folder_name}')

        file_name = file_name.split('.')[0]
        sheet = file_name.split('_')[-1]
        file_name_for_print = file_name.split('_')[0]
        print(f"{file_name_for_print} Sheet:{sheet} similarity", score * 100)

        # cv2.imwrite(f'{path_to_result}{folder_name}/{file_name}_before_conv.png', before)
        # cv2.imwrite(f'{path_to_result}{folder_name}/{file_name}_after_conv.png', after)
        # cv2.imwrite(os.getcwd() + f'\\data/{number_of_pages}diff.png', diff)
        # cv2.imwrite(os.getcwd() + f'\\data/{number_of_pages}_mask.png', mask)
        # cv2.imwrite(f'{path_to_result}{folder_name}/{file_name}_filled_after.png', filled_after)
        # images = [imageio.imread(f'{path_to_result}{folder_name}/{file_name}_before_conv.png'),
        #           imageio.imread(f'{path_to_result}{folder_name}/{file_name}_after_conv.png')]
        images = [before,
                  after]
        imageio.mimsave(f'{path_to_result}{folder_name}/{file_name}.gif', images, duration=1)
        # Helper.delete(f'{path_to_result}{folder_name}/{file_name}_before_conv.png')
        # Helper.delete(f'{path_to_result}{folder_name}/{file_name}_after_conv.png')


    @staticmethod
    def grab(path, filename, number_of_pages):
        sleep(wait_for_open)
        extension = filename.split('.')[-1]
        filename_for_screen = filename.replace(f'.{extension}', '')
        num_of_sheet = 1
        for el in range(int(number_of_pages)):
            im = ImageGrab.grab()
            im.save(path + filename_for_screen + str(f'_{num_of_sheet}') + '.png', 'PNG')
            # grab(filename + str(number_of_pages), path2)
            pg.click()
            pg.press('pgdn')
            sleep(wait_for_press)
            num_of_sheet += 1

    # def open_document_for_compare(list_of_files, from_extension, to_extension):
    #     for file_name in list_of_files:
    #         extension = file_name.split('.')[-1]
    #         if extension == 'ppt' or extension == 'pptx':
    #             run(file_name, 'POWERPNT.EXE')
    #
    #         elif extension == 'doc' or 'docx':
    #             word, statistics_word = word_opener(file_name, custom_path_to_document_to,
    #                                                 path_to_tmpimg_after_conversion)
    #             grab(path_to_tmpimg_after_conversion, file_name.split('.')[0], statistics_word['num_of_sheets'])
    #             word.Close()
    #             file_name_from = file_name.replace(f'.{to_extension}', f'.{from_extension}')
    #             word, statistics_word = word_opener(file_name_from, custom_path_to_document_from,
    #                                                 path_to_tmpimg_befor_conversion)
    #             word.Close()
    # sb.call(["taskkill", "/IM", "WINWORD.EXE"])
    # sb.call(["taskkill", "/IM", "WINWORD.EXE"])
    # sb.call(f'powershell.exe kill -Name WINWORD', shell=True)

    def start_to_compare_images(self):
        list_of_images = os.listdir(path_to_tmpimg_after_conversion)
        for image_name in track(list_of_images, description='Comparing In Progress'):

            if os.path.exists(f'{path_to_tmpimg_befor_conversion}{image_name}') and os.path.exists(
                    f'{path_to_tmpimg_after_conversion}{image_name}'):
                self.compare_img(path_to_tmpimg_befor_conversion + image_name,
                                 path_to_tmpimg_after_conversion + image_name,
                                 image_name)
            else:
                print(f'[bold red] FIle not found [/bold red]{image_name}')

        # Helper.delete(f'{path_to_tmpimg_after_conversion}*')
        # Helper.delete(f'{path_to_tmpimg_befor_conversion}*')

    # def run_compare_statistics():
    #     a = json.load(open(path_to_result + 'Alimentaire_Etude_Planning_StrategiqueKM_2010_doc.json'))
    #     b = json.load(open(path_to_result + 'Alimentaire_Etude_Planning_StrategiqueKM_2010_docx.json'))
    #
    #     modified = dict_compare(a, b)
    #     print(modified)
    #

# if __name__ == "__main__":
#     # sb.call(f'powershell.exe rm {path_to_tmpimg_befor_conversion}*,'
#         f'{path_to_tmpimg_after_conversion}* -Recurse',
#         shell=True)
#
# create_project_dirs()
# open_document_for_compare(list_file_names_doc_from_compare, extension_from, extension_to)

# start_to_compare_images()
# sb.call(f'powershell.exe rm {path_to_tmpimg_befor_conversion}*,'
#         f'{path_to_tmpimg_after_conversion}* -Recurse',
#         shell=True)

# list_of_files = os.listdir(path_to_compare_files)
