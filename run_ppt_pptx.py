from tqdm import tqdm

from config import *
from libs.functional.presentation.ppt_to_pptx_compare import PowerPoint

if __name__ == "__main__":
    for i in tqdm(range(1)):
        # PowerPoint(list_of_file_names)
        PowerPoint(os.listdir(converted_doc_folder))
