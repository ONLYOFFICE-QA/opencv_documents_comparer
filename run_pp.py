from tqdm import tqdm

from libs.power_point import PowerPoint
from var import *

if __name__ == "__main__":
    for i in tqdm(range(1)):
        PowerPoint(list_file_names_doc_from_compare)
        # PowerPoint(os.listdir(custom_doc_to))
