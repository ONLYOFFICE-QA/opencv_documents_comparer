import cv2
import imageio
import numpy as np
from PIL import ImageGrab
from skimage.metrics import structural_similarity

from libs.helper import *
from var import *


class CompareImage(Helper):
    def __init__(self, file_name):
        self.start_to_compare_images(file_name)

    def compare_img(self, before_conversion, after_conversion, image_name, file_name):
        before = cv2.imread(before_conversion)
        after = cv2.imread(after_conversion)
        before_gray = cv2.cvtColor(before, cv2.COLOR_BGR2GRAY)
        after_gray = cv2.cvtColor(after, cv2.COLOR_BGR2GRAY)

        (score, diff) = structural_similarity(before_gray, after_gray, full=True)

        diff = (diff * 255).astype("uint8")

        thresh = cv2.threshold(diff, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)[1]
        contours = cv2.findContours(thresh.copy(), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        contours = contours[0] if len(contours) == 2 else contours[1]
        color = (0, 0, 255)
        color_gif = (255, 0, 0)
        for c in contours:
            area = cv2.contourArea(c)
            if area > 40:
                x, y, w, h = cv2.boundingRect(c)
                cv2.rectangle(before, (x, y), (x + w, y + h), color, 0)
                cv2.rectangle(after, (x, y), (x + w, y + h), color, 0)

        folder_name = image_name.replace(f'_{image_name.split("_")[-1]}', '')
        screen_folder = 'screen'

        Helper.create_dir(f'{path_to_result}{folder_name}')
        Helper.create_dir(f'{path_to_result}{folder_name}/{screen_folder}')

        similarity = score * 100
        image_name = image_name.split('.')[0]
        sheet = image_name.split('_')[-1]
        file_name_for_print = image_name.split('_')[0]
        print(f"{file_name_for_print} Sheet:{sheet} similarity", similarity)

        cv2.putText(before, f'Before sheet {sheet}', (50, 300), cv2.FONT_HERSHEY_COMPLEX, 1, color=color,
                    thickness=2)
        cv2.putText(after, f'After sheet {sheet}', (50, 300), cv2.FONT_HERSHEY_COMPLEX, 1, color=color,
                    thickness=2)
        cv2.putText(after, f'Similarity {round(similarity, 3)}%', (50, 400), cv2.FONT_HERSHEY_COMPLEX, 1,
                    color=color,
                    thickness=2)
        cv2.putText(before, f'Similarity {round(similarity, 3)}%', (50, 400), cv2.FONT_HERSHEY_COMPLEX, 1,
                    color=color,
                    thickness=2)

        images = [before,
                  after]
        collage = np.hstack([before, after])

        file_name_for_save = f'{image_name}_similarity.gif'
        file_name_from = file_name.replace(f'.{extension_to}', f'.{extension_from}')
        if similarity < 98:
            self.create_dir(f'{path_to_errors_sim_file}{folder_name}')
            Helper.create_dir(f'{path_to_errors_sim_file}{folder_name}/{screen_folder}')
            self.copy_to_errors_sim(file_name, file_name_from)
            cv2.imwrite(f'{path_to_errors_sim_file}{folder_name}/{screen_folder}/{file_name}_{sheet}_collage.png',
                        collage)
            imageio.mimsave(f'{path_to_errors_sim_file}{folder_name}/{file_name_for_save}',
                            images,
                            duration=1)
        else:
            imageio.mimsave(f'{path_to_result}{folder_name}/{file_name_for_save}',
                            images,
                            duration=1)
            cv2.imwrite(f'{path_to_errors_sim_file}{folder_name}/{screen_folder}/{file_name}_{sheet}_collage.png',
                        collage)

    @staticmethod
    def grab_coordinate(path, filename, number_of_pages, coordinate):
        img_name = filename.replace(f'.{filename.split(".")[-1]}', '')

        im = ImageGrab.grab(bbox=(coordinate[0],
                                  coordinate[1] + 170,
                                  coordinate[2] - 30,
                                  coordinate[3] - 20))

        im.save(path + img_name + str(f'_{number_of_pages}') + '.png', 'PNG')

    @staticmethod
    def grab_coordinate_pp(path, filename, number_of_pages, coordinate):
        img_name = filename.replace(f'.{filename.split(".")[-1]}', '')

        im = ImageGrab.grab(bbox=(coordinate[0] + 350,
                                  coordinate[1] + 170,
                                  coordinate[2] - 50,
                                  coordinate[3] - 20))

        im.save(path + img_name + str(f'_{number_of_pages}') + '.png', 'PNG')

    @staticmethod
    def grab(path, filename, number_of_pages):
        filename_for_screen = filename.replace(f'.{filename.split(".")[-1]}', '')
        im = ImageGrab.grab()
        im.save(path + filename_for_screen + str(f'_{number_of_pages}') + '.png', 'PNG')

    def start_to_compare_images(self, file_name):
        list_of_images = os.listdir(tmp_after)
        for image_name in track(list_of_images, description='Comparing In Progress'):

            if os.path.exists(f'{tmp_befor}{image_name}') \
                    and os.path.exists(f'{tmp_after}{image_name}'):

                self.compare_img(tmp_befor + image_name,
                                 tmp_after + image_name,
                                 image_name,
                                 file_name)

            else:
                print(f'[bold red] File not found [/bold red]{image_name}')

        # Helper.delete(f'{tmp_after}*')
        # Helper.delete(f'{tmp_befor}*')
