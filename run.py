from libs.Compare_Image import CompareImage
from libs.Word_image_compare import WordCompareImg
from var import *
from tqdm import tqdm
from multiprocessing import Process


from libs.get_error import run_get_pr
from libs.word import Word


if __name__ == "__main__":
    for i in tqdm(range(1)):
        p1 = Process(target=run_get_pr)
        p1.start()
        # WordCompareImg()
        # CompareImage()
        Word(os.listdir(custom_doc_to))
        # Word(list_file_names_doc_from_compare)
        p1.terminate()

