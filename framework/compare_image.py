# -*- coding: utf-8 -*-
import os
import re
from os.path import join

import cv2
import imageio
import numpy as np
from PIL import ImageGrab
from loguru import logger
from rich import print
from rich.progress import track
from skimage.metrics import structural_similarity

from framework.StaticData import StaticData
from framework.FileUtils import FileUtils


class CompareImage:
    def __init__(self, coefficient=98):
        self.coefficient = coefficient
        self.helper = StaticData.DOC_ACTIONS
        self.similarity: int = 0
        self.image_name, self.sheet_num, self.gif_name = '', '', ''
        self.source_img, self.converted_img, self.src_processed_img, self.cnv_processed_img = [], [], [], []
        self.result_folder_name = self.get_result_folder_name()
        self.img_comparison_diff_dir = join(self.helper.result_folder, 'img_comparison_diff', self.result_folder_name)
        self.passed = join(self.helper.result_folder, 'passed')
        self.start_compare_images()

    def get_result_folder_name(self):
        result_folder_name = self.helper.converted_file.replace(f'.{self.helper.converted_extension}', '')
        result_folder_name = re.sub(r"^\s+|\n|\r|\.|,|-|\s|\s+$", '', result_folder_name)
        result_folder_name = result_folder_name[:35]
        return result_folder_name

    def start_compare_images(self):
        description = "[bold blue]Comparing In Progress[/]"
        for self.image_name in track(os.listdir(StaticData.TMP_DIR_CONVERTED_IMG), description=description):
            if os.path.exists(join(StaticData.TMP_DIR_SOURCE_IMG, self.image_name)) \
                    and os.path.exists(join(StaticData.TMP_DIR_CONVERTED_IMG, self.image_name)):
                self.sheet_num = self.image_name.split('_')[-1].replace('.png', '')
                self.gif_name = f"{self.image_name.split('.')[0]}_similarity.gif"
                self.source_img = cv2.imread(join(StaticData.TMP_DIR_SOURCE_IMG, self.image_name))
                self.converted_img = cv2.imread(join(StaticData.TMP_DIR_CONVERTED_IMG, self.image_name))
                self.image_handler()
                print(f"[blue]{self.helper.converted_file} Sheet: {self.sheet_num} similarity[/]: {self.similarity}")
                self.save_result() if self.similarity < self.coefficient else print('[bold green]passed')
            else:
                logger.error(f'Image {self.image_name} not found, copied file to "Untested"')
                self.helper.copy_testing_files_to_folder(self.helper.untested_folder)
        self.helper.tmp_cleaner()

    def put_information_on_img(self):
        self.put_text(self.converted_img, f'After')
        self.put_text(self.source_img, f'Before')
        self.put_text(self.src_processed_img, f'Before sheet: {self.sheet_num}. Similarity{self.similarity}%')
        self.put_text(self.cnv_processed_img, f'After sheet: {self.sheet_num}. Similarity{self.similarity}%')

    def image_handler(self):
        if self.helper.converted_file.lower().endswith(".xlsx"):
            self.find_difference(self.source_img, self.converted_img)
        else:
            try:
                self.find_difference(self.find_contours(self.source_img), self.find_contours(self.converted_img))
            except Exception:
                self.find_difference(self.source_img, self.converted_img)
        self.put_information_on_img()

    def save_result(self):
        self.helper.copy_testing_files_to_folder(self.img_comparison_diff_dir)
        FileUtils.create_dir(join(self.img_comparison_diff_dir, 'gif'), silence=True)
        FileUtils.create_dir(join(self.img_comparison_diff_dir, 'screen'), silence=True)
        collage = np.hstack([self.source_img, self.converted_img])
        cv2.imwrite(join(self.img_comparison_diff_dir, 'screen', f"{self.image_name}_collage.png"), collage)
        images = [self.src_processed_img, self.cnv_processed_img]
        imageio.mimsave(join(self.img_comparison_diff_dir, 'gif', self.gif_name), images, duration=1)

    def find_contours(self, img):
        rgb, gray = cv2.cvtColor(img, cv2.COLOR_BGR2RGB), cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        _, binary = cv2.threshold(gray, 125, 255, cv2.THRESH_BINARY)
        contours, hierarchy = cv2.findContours(binary, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        for contour in contours:
            x, y, w, h = cv2.boundingRect(contour)
            if h >= 500:
                return rgb[y:y + h + 1, x:x + w + 1] if self.helper.source_extension == 'odp' else rgb[y:y + h, x:x + w]

    def find_difference(self, src_img, cnv_img):
        before_gray, after_gray = cv2.cvtColor(src_img, cv2.COLOR_BGR2GRAY), cv2.cvtColor(cnv_img, cv2.COLOR_BGR2GRAY)
        (score, diff) = structural_similarity(before_gray, after_gray, full=True)
        diff = (diff * 255).astype("uint8")
        thresh = cv2.threshold(diff, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)[1]
        contours = cv2.findContours(thresh.copy(), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        contours = contours[0] if len(contours) == 2 else contours[1]
        for contur in contours:
            area = cv2.contourArea(contur)
            if area > 40:
                x, y, w, h = cv2.boundingRect(contur)
                cv2.rectangle(src_img, (x, y), (x + w, y + h), (0, 0, 255), 0)
                cv2.rectangle(cnv_img, (x, y), (x + w, y + h), (0, 0, 255), 0)
        self.similarity = round(score * 100, 3)
        self.src_processed_img, self.cnv_processed_img = src_img, cnv_img

    @staticmethod
    def put_text(cv2_opened_image, text):
        cv2.putText(cv2_opened_image, text, (20, 35), cv2.FONT_HERSHEY_COMPLEX, 1, color=(0, 0, 255), thickness=2)

    @staticmethod
    def grab_coordinate(path_to_img_for_save, coordinate):
        image = ImageGrab.grab(bbox=coordinate)
        image.save(path_to_img_for_save, 'PNG')
