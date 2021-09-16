from multiprocessing import Process

from tqdm import tqdm

from libs.Compare_Image import CompareImage
from libs.Exel_image_compaire import ExelCompareImage
from libs.get_error import run_get_error_exel
from var import *

if __name__ == "__main__":
    for i in tqdm(range(1)):
        p1 = Process(target=run_get_error_exel)
        p1.start()
        # ExelCompareImage(list_file_names_doc_from_compare)
        ExelCompareImage(os.listdir(r'd:\Users\kokol\Desktop\12123'))
        # Exel(os.listdir(custom_doc_to))
        CompareImage('kef.docx', 100)
        p1.terminate()
