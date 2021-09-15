from multiprocessing import Process

from tqdm import tqdm

from libs.Exel import Exel
from libs.get_error import run_get_error_exel
from var import *

if __name__ == "__main__":
    for i in tqdm(range(1)):
        p1 = Process(target=run_get_error_exel)
        p1.start()
        Exel(list_file_names_doc_from_compare)
        # Exel(os.listdir(custom_doc_to))
        # CompareImage('kef.docx')
        p1.terminate()
