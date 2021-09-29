from tqdm import tqdm

from libs.functional.presentation.ppt_to_pptx_compare import PowerPoint
from var import *

if __name__ == "__main__":
    for i in tqdm(range(1)):
        # PowerPoint(list_of_file_names)
        PowerPoint(os.listdir(converted_doc_folder))
