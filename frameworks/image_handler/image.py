# -*- coding: utf-8 -*-
import cv2
import imageio
import numpy as np
import mss
import mss.tools

from skimage.metrics import structural_similarity


class Image:
    """
    Utility class for working with images, including reading, finding templates, checking presence, and more.
    """

    @staticmethod
    def read(img_path: str) -> cv2.imread:
        """
        Reads an image from the specified path.
        :param img_path: Path to the image file.
        :return: Loaded image as a NumPy array.
        """
        return cv2.imread(img_path)

    @staticmethod
    def find_template_on_window(
            window_coord: tuple,
            template: str,
            threshold: int | float = 0.8
    ) -> list[int, int] | None:
        """
        Finds a template image within a window.
        :param window_coord: Coordinates of the window to search within.
        :param template: Path to the template image file.
        :param threshold: Threshold for template matching.
        :return: Coordinates of the found template's center or None if not found.
        """
        window = cv2.cvtColor(Image.grab_coordinate(window_coord), cv2.COLOR_BGR2GRAY)
        template = cv2.cvtColor(cv2.imread(template), cv2.COLOR_BGR2GRAY)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(cv2.matchTemplate(window, template, cv2.TM_CCOEFF_NORMED))
        if max_val >= threshold:
            h, w = template.shape
            center_x = max_loc[0] + w // 2 + window_coord[0]
            center_y = max_loc[1] + h // 2 + window_coord[1]
            return [center_x, center_y]
        return None

    @staticmethod
    def is_image_present(window_coordinates: tuple, template_path: str, threshold: int | float = 0.8) -> bool:
        """
        Checks if an image is present within a specified window.
        :param window_coordinates: Coordinates of the window to search within.
        :param template_path: Path to the template image file.
        :param threshold: Threshold for template matching.
        :return: True if the image is present, False otherwise.
        """
        window = cv2.cvtColor(Image.grab_coordinate(window_coordinates), cv2.COLOR_BGR2GRAY)
        template = cv2.cvtColor(cv2.imread(template_path), cv2.COLOR_BGR2GRAY)
        _, max_val, _, _ = cv2.minMaxLoc(cv2.matchTemplate(window, template, cv2.TM_CCOEFF_NORMED))
        return True if max_val >= threshold else False

    @staticmethod
    def grab_coordinate(window_coordinates: tuple) -> np.array:
        """
        Grabs a screenshot of the specified window coordinates.
        :param window_coordinates: Coordinates of the window to grab.
        :return: Screenshot as a NumPy array.
        """
        left, top, right, bottom = window_coordinates
        with mss.mss() as sct:
            return np.array(sct.grab({"left": left, "top": top, "width": right - left, "height": bottom - top}))

    @staticmethod
    def find_contours(img):
        """
        Finds contours within an image.
        :param img: Input image.
        :return: Image with contours drawn.
        """
        rgb, gray = cv2.cvtColor(img, cv2.COLOR_BGR2RGB), cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        _, binary = cv2.threshold(gray, 125, 255, cv2.THRESH_BINARY)
        contours, hierarchy = cv2.findContours(binary, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        for contour in contours:
            x, y, w, h = cv2.boundingRect(contour)
            if h >= 500:
                # TODO fix it for odp: return rgb[y:y + h + 1, x:x + w + 1] if source_ext == 'odp' else rgb[y:y + h,
                #  x:x + w]
                return rgb[y:y + h, x:x + w]

    @staticmethod
    def find_difference(img_1: np.ndarray, img_2: np.ndarray) -> tuple:
        """
        Finds the structural similarity difference between two images.
        :param img_1: First image.
        :param img_2: Second image.
        :return: Structural similarity difference.
        """
        before = cv2.cvtColor(img_1, cv2.COLOR_BGR2GRAY)
        after = cv2.cvtColor(img_2, cv2.COLOR_BGR2GRAY)

        if before.shape != after.shape:
            before, after = Image.align_sizes(before, after)

        similarity, difference = structural_similarity(before, after, full=True)
        return similarity, difference

    @staticmethod
    def align_sizes(img1, img2):

        height_diff = img1.shape[0] - img2.shape[0]
        width_diff = img1.shape[1] - img2.shape[1]

        if height_diff > 0:
            img2 = cv2.copyMakeBorder(img2, 0, height_diff, 0, 0, cv2.BORDER_CONSTANT, value=(0, 0, 0))
        elif height_diff < 0:
            img1 = cv2.copyMakeBorder(img1, 0, -height_diff, 0, 0, cv2.BORDER_CONSTANT, value=(0, 0, 0))

        if width_diff > 0:
            img2 = cv2.copyMakeBorder(img2, 0, 0, 0, width_diff, cv2.BORDER_CONSTANT, value=(0, 0, 0))
        elif width_diff < 0:
            img1 = cv2.copyMakeBorder(img1, 0, 0, 0, -width_diff, cv2.BORDER_CONSTANT, value=(0, 0, 0))

        return img1, img2

    @staticmethod
    def show(image) -> None:
        cv2.imshow('image', image)
        cv2.waitKey(0)
        cv2.destroyAllWindows()

    @staticmethod
    def draw_differences(img_1: np.ndarray, img_2: np.ndarray, diff: np.ndarray) -> tuple[np.ndarray, np.ndarray]:
        """
        Draws differences between two images.
        :param img_1: First image.
        :param img_2: Second image.
        :param diff: Difference image.
        :return: Tuple containing the first and second images with differences drawn.
        """
        thresh = cv2.threshold((diff * 255).astype("uint8"), 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)[1]
        contours = cv2.findContours(thresh.copy(), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        for contur in contours[0] if len(contours) == 2 else contours[1]:
            if cv2.contourArea(contur) > 40:
                x, y, w, h = cv2.boundingRect(contur)
                cv2.rectangle(img_1, (x, y), (x + w, y + h), (0, 0, 255), 0)
                cv2.rectangle(img_2, (x, y), (x + w, y + h), (0, 0, 255), 0)
        return img_1, img_2

    @staticmethod
    def save(path, img):
        """
        Saves an image to the specified path.
        :param path: Path to save the image.
        :param img: Image to save.
        :return: None
        """
        cv2.imwrite(path, img)

    @staticmethod
    def save_gif(save_path: str, img_paths: list, duration: int = 1000, loop: int = 0):
        """
        Saves a list of images as a GIF.
        :param save_path: Path to save the GIF.
        :param img_paths: List of image paths.
        :param duration: Duration of each frame in milliseconds.
        :param loop: number of repetitions, 0 - repeat indefinitely
        :return: None
        """
        imageio.mimsave(save_path, img_paths, duration=duration, loop=loop)

    @staticmethod
    def put_text(cv2_opened_image, text: str) -> None:
        """
        Adds text to an image.
        :param cv2_opened_image: Opened image using cv2.
        :param text: Text to add.
        :return: None
        """
        cv2.putText(cv2_opened_image, text, (20, 35), cv2.FONT_HERSHEY_COMPLEX, 1, color=(0, 0, 255), thickness=2)

    @staticmethod
    def make_screenshot(img_path: str, coordinate: tuple) -> None:
        """
        Makes a screenshot of a specified coordinate region.
        :param img_path: Path to save the screenshot.
        :param coordinate: Coordinates of the region to screenshot.
        :return: None
        """
        left, top, right, bottom = coordinate
        with mss.mss() as sct:
            img = sct.grab({"left": left, "top": top, "width": right - left, "height": bottom - top})
            mss.tools.to_png(img.rgb, img.size, output=img_path)
