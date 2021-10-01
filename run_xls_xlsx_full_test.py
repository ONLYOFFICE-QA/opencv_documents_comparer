from multiprocessing import Process

from tqdm import tqdm

from config import *
from libs.functional.spreadsheets.xls_to_xlsx_image_compare import ExcelCompareImage
from libs.functional.spreadsheets.xls_to_xlsx_statistic_compare import Excel
from libs.helpers.get_error import run_get_error_exel

if __name__ == "__main__":
    for i in tqdm(range(1)):
        error_processing = Process(target=run_get_error_exel)
        error_processing.start()
        Excel(os.listdir(converted_doc_folder))
        ExcelCompareImage(differences_statistic)
        error_processing.terminate()
