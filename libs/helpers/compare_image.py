import cv2
import imageio
import numpy as np
from PIL import ImageGrab
from rich.progress import track
from skimage.metrics import structural_similarity

from libs.helpers.helper import *
from var import *


class CompareImage(Helper):
    def __init__(self, file_name, koff=98):
        self.start_to_compare_images(file_name, koff)

    def compare_img(self, before_conversion, after_conversion, image_name, converted_file, koff):
        folder_name = converted_file.replace(f'.{converted_file.split(".")[-1]}', '')
        screen_folder = 'screen'
        Helper.create_dir(f'{passed}{folder_name}')
        Helper.create_dir(f'{passed}{folder_name}/{screen_folder}')
        image_name = image_name.split('.')[0]
        sheet = image_name.split('_')[-1]

        before = cv2.imread(before_conversion)
        after = cv2.imread(after_conversion)
        before_gray = cv2.cvtColor(before, cv2.COLOR_BGR2GRAY)
        after_gray = cv2.cvtColor(after, cv2.COLOR_BGR2GRAY)

        after_for_collage = self.put_text(after, f'After')
        before_for_collage = self.put_text(before, f'Before')
        collage = np.hstack([before_for_collage, after_for_collage])

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

        print(f"{converted_file} Sheet: {sheet} similarity: {similarity}")

        before = self.put_text(before, f'Before sheet {sheet}. Similarity {round(similarity, 3)}%')
        after = self.put_text(after, f'After sheet {sheet}. Similarity {round(similarity, 3)}%')

        images = [before, after]

        file_name_for_gif = f'{image_name}_similarity.gif'
        source_file = converted_file.replace(f'.{extension_converted}', f'.{extension_source}')
        if similarity < koff:
            self.create_dir(f'{differences_compare_image}{folder_name}')
            self.create_dir(f'{differences_compare_image}{folder_name}/gif/')
            self.create_dir(f'{differences_compare_image}{folder_name}/{screen_folder}')
            self.copy_to_folder(converted_file, source_file, f'{differences_compare_image}{folder_name}/')
            cv2.imwrite(
                f'{differences_compare_image}{folder_name}/{screen_folder}/{converted_file}_{sheet}_collage.png',
                collage)
            imageio.mimsave(f'{differences_compare_image}{folder_name}/gif/{file_name_for_gif}',
                            images,
                            duration=1)
        else:
            imageio.mimsave(f'{passed}{folder_name}/{file_name_for_gif}',
                            images,
                            duration=1)
            cv2.imwrite(
                f'{differences_compare_image}{folder_name}/{screen_folder}/{converted_file}_{sheet}_collage.png',
                collage)

    @staticmethod
    def grab_coordinate_exel(path, filename, list_num, number_of_pages, coordinate):
        img_name = filename.replace(f'.{filename.split(".")[-1]}', '')

        image = ImageGrab.grab(bbox=coordinate)

        image.save(path + img_name + str(f'_list_{list_num}_page_{number_of_pages}') + '.png', 'PNG')

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
        pass

    @staticmethod
    def grab_coordinate(path, filename, number_of_pages, coordinate):
        img_name = filename.replace(f'.{filename.split(".")[-1]}', '')

        image = ImageGrab.grab(bbox=coordinate)

        image.save(path + img_name + str(f'_{number_of_pages}') + '.png', 'PNG')

    @staticmethod
    def grab(path, filename, number_of_pages):
        filename_for_screen = filename.replace(f'.{filename.split(".")[-1]}', '')
        im = ImageGrab.grab()
        im.save(path + filename_for_screen + str(f'_{number_of_pages}') + '.png', 'PNG')

    def start_to_compare_images(self, file_name, koff):
        list_of_images = os.listdir(tmp_converted_image)
        for image_name in track(list_of_images, description='Comparing In Progress'):

            if os.path.exists(f'{tmp_source_image}{image_name}') \
                    and os.path.exists(f'{tmp_converted_image}{image_name}'):

                self.compare_img(f'{tmp_source_image}{image_name}',
                                 f'{tmp_converted_image}{image_name}',
                                 image_name,
                                 file_name,
                                 koff)

            else:
                print(f'[bold red] File not found [/bold red]{image_name}')

        Helper.delete(f'{tmp_converted_image}*')
        Helper.delete(f'{tmp_source_image}*')
