import subprocess as sb
from multiprocessing import Process

from tqdm import tqdm

from libs.functional.documents.doc_to_docx_statistic_compare import Word
from libs.helpers.get_error import run_get_errors_pp
from var import *

if __name__ == "__main__":
    for i in tqdm(range(1)):
        p1 = Process(target=run_get_errors_pp)
        p1.start()
        sb.call(f'powershell.exe kill -Name WINWORD', shell=True)
        Word(os.listdir(converted_doc_folder))
        # Word(list_of_file_names)
        p1.terminate()
