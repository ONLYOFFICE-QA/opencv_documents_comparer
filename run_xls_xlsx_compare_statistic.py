from multiprocessing import Process

from tqdm import tqdm

from libs.functional.spreadsheets.xls_to_xlsx_statistic_compare import Excel
from libs.helpers.get_error import run_get_error_exel
from variables import *

if __name__ == "__main__":
    for i in tqdm(range(1)):
        error_processing = Process(target=run_get_error_exel)
        error_processing.start()
        Excel(os.listdir(converted_doc_folder))
        error_processing.terminate()
