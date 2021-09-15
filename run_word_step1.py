import subprocess as sb
from multiprocessing import Process

from tqdm import tqdm

from libs.get_error import run_get_pr
from libs.word import Word
from var import *

if __name__ == "__main__":
    for i in tqdm(range(1)):
        p1 = Process(target=run_get_pr)
        p1.start()
        sb.call(f'powershell.exe kill -Name WINWORD', shell=True)
        Word(os.listdir(custom_doc_to))
        # Word(list_file_names_doc_from_compare)
        p1.terminate()
