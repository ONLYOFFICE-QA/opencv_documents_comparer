# -*- coding: utf-8 -*-
import re

import cv2
import imageio
import numpy as np
from PIL import ImageGrab
from rich.progress import track
from skimage.metrics import structural_similarity

from libs.helpers.helper import *


class CompareImage:

    def __init__(self, helper, koff=98, odp=False):
        self.odp = odp
        self.koff = koff
        self.helper = helper
        self.screen_folder = 'screen'
        self.folder_name = self.helper.converted_file.replace(f'.{self.helper.converted_file.split(".")[-1]}', '')
        self.folder_name = re.sub(r"^\s+|\n|\r|\.|,|-|\s|\s+$", '', self.folder_name)
        self.folder_name = self.folder_name[:35]
        Helper.create_dir(f'{self.helper.passed}{self.folder_name}')
        Helper.create_dir(f'{self.helper.passed}{self.folder_name}/{self.screen_folder}')
        self.start_to_compare_images()

    def start_to_compare_images(self):
        for image_name in track(os.listdir(self.helper.tmp_dir_converted_image),
                                description='[bold blue]Comparing In Progress[/bold blue]'):

            if os.path.exists(f'{self.helper.tmp_dir_source_image}{image_name}') \
                    and os.path.exists(f'{self.helper.tmp_dir_converted_image}{image_name}'):

                self.compare_img(f'{self.helper.tmp_dir_source_image}{image_name}',
                                 f'{self.helper.tmp_dir_converted_image}{image_name}',
                                 image_name,
                                 )

            else:
                logger.error(f'Image {image_name} not found, '
                             f'copied file {self.helper.converted_file} '
                             f'and {self.helper.source_file} to "Untested"')

                self.helper.copy_to_folder(self.helper.untested_folder)

        self.helper.delete(f'{self.helper.tmp_dir_converted_image}*')
        self.helper.delete(f'{self.helper.tmp_dir_source_image}*')

    def compare_img(self, image_before_conversion, image_after_conversion, image_name):
        image_name = image_name.split('.')[0]
        sheet = image_name.split('_')[-1]
        file_name_for_gif = f'{image_name}_similarity.gif'
        before_full = cv2.imread(image_before_conversion)
        after_full = cv2.imread(image_after_conversion)

        if self.helper.converted_file.endswith((".xlsx", ".XLSX")):
            after, before, similarity = self.find_difference(after_full, before_full)
            pass
        else:
            if self.odp:
                before = self.find_contours(image_before_conversion, odp=True)
            else:
                before = self.find_contours(image_before_conversion)
            after = self.find_contours(image_after_conversion)

            try:
                after, before, similarity = self.find_difference(after, before)
            except Exception:
                after, before, similarity = self.find_difference(after_full, before_full)
                pass

        print(f"[bold blue]{self.helper.converted_file}[/bold blue] Sheet: {sheet} "
              f"[bold blue]similarity[/bold blue]: {similarity}")

        before = self.put_text(before, f'Before sheet {sheet}. Similarity {round(similarity, 3)}%')
        after = self.put_text(after, f'After sheet {sheet}. Similarity {round(similarity, 3)}%')
        collage = self.collage(after_full, before_full)
        images = [before, after]
        if similarity < self.koff:
            self.write_down_results(collage, images, file_name_for_gif, image_name)
        else:
            print('[bold green]passed[/bold green]')
            imageio.mimsave(f'{self.helper.passed}{self.folder_name}/{file_name_for_gif}',
                            images,
                            duration=1)
            cv2.imwrite(
                f'{self.helper.passed}{self.folder_name}/{self.screen_folder}/{image_name}_collage.png',
                collage)

    def collage(self, after_img, before_img):
        after_for_collage = self.put_text(after_img, f'After')
        before_for_collage = self.put_text(before_img, f'Before')
        collage = np.hstack([before_for_collage, after_for_collage])
        return collage

    def write_down_results(self, collage, images, file_name_for_gif, image_name):
        self.helper.create_dir(f'{self.helper.differences_compare_image}{self.folder_name}')
        self.helper.create_dir(f'{self.helper.differences_compare_image}{self.folder_name}/gif/')
        self.helper.create_dir(f'{self.helper.differences_compare_image}{self.folder_name}/{self.screen_folder}')
        self.helper.copy_to_folder(f'{self.helper.differences_compare_image}{self.folder_name}/')
        cv2.imwrite(
            f'{self.helper.differences_compare_image}{self.folder_name}/{self.screen_folder}/'
            f'{image_name}_collage.png',
            collage)
        imageio.mimsave(f'{self.helper.differences_compare_image}{self.folder_name}/gif/{file_name_for_gif}',
                        images,
                        duration=1)
        pass

    @staticmethod
    def find_contours(img, odp=False):
        img_source = cv2.imread(img)
        img = cv2.cvtColor(img_source, cv2.COLOR_BGR2RGB)
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        _, binary = cv2.threshold(gray, 125, 255, cv2.THRESH_BINARY)
        contours, hierarchy = cv2.findContours(binary, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        for c in contours:
            x, y, w, h = cv2.boundingRect(c)
            if h >= 500:
                # to save the images
                if odp:
                    img = img[y:y + h + 1, x:x + w + 1]
                else:
                    img = img[y:y + h, x:x + w]
                return img

    @staticmethod
    def find_difference(after, before):
        before_gray = cv2.cvtColor(before, cv2.COLOR_BGR2GRAY)
        after_gray = cv2.cvtColor(after, cv2.COLOR_BGR2GRAY)
        (score, diff) = structural_similarity(before_gray, after_gray, full=True)
        diff = (diff * 255).astype("uint8")
        thresh = cv2.threshold(diff, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)[1]
        contours = cv2.findContours(thresh.copy(), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        contours = contours[0] if len(contours) == 2 else contours[1]
        color = (0, 0, 255)
        for c in contours:
            area = cv2.contourArea(c)
            if area > 40:
                x, y, w, h = cv2.boundingRect(c)
                cv2.rectangle(before, (x, y), (x + w, y + h), color, 0)
                cv2.rectangle(after, (x, y), (x + w, y + h), color, 0)
        similarity = score * 100
        return after, before, similarity

    @staticmethod
    def put_text(image, text):
        color = (0, 0, 255)
        cv2.putText(image, f'{text}',
                    (20, 35),
                    cv2.FONT_HERSHEY_COMPLEX,
                    1,
                    color=color,
                    thickness=2)
        return image

    @staticmethod
    def grab_coordinate_exel(path, list_num, number_of_pages, coordinate):
        image = ImageGrab.grab(bbox=coordinate)
        image.save(path + str(f'_list_{list_num}_page_{number_of_pages}') + '.png', 'PNG')

    @staticmethod
    def grab_coordinate(path, number_of_pages, coordinate):
        image = ImageGrab.grab(bbox=coordinate)
        image.save(path + str(f'page_{number_of_pages}') + '.png', 'PNG')

    @staticmethod
    def grab(path, filename, number_of_pages):
        filename_for_screen = filename.replace(f'.{filename.split(".")[-1]}', '')
        im = ImageGrab.grab()
        im.save(path + filename_for_screen + str(f'_{number_of_pages}') + '.png', 'PNG')
