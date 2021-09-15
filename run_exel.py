from tqdm import tqdm

from libs.Exel import Exel
from var import *

if __name__ == "__main__":
    for i in tqdm(range(1)):
        Exel(list_file_names_doc_from_compare)
        # CompareImage('kef.docx')
