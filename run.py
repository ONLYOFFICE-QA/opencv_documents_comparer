from tqdm import tqdm

from libs.power_point import PowerPoint
from var import *

if __name__ == "__main__":
    for i in tqdm(range(1)):
        # p1 = Process(target=run_get_pr)
        # p1.start()
        # PowerPoint(list_file_names_doc_from_compare)
        PowerPoint(os.listdir(custom_doc_to))

        # print(os.listdir('d:\\Users\\kokol\\Desktop\\data\\6.4.1.8_odt_docx\\errors'))
        # WordCompareImg()
        # sb.call(f'powershell.exe kill -Name WINWORD', shell=True)
        # CompareImage()
        # Word(os.listdir(custom_doc_to))
        # Word(list_file_names_doc_from_compare)
        # p1.terminate()
        # subprocess.call(["TASKKILL", "/IM", "POWERPNT.EXE", "/t", "/f"], shell=True)
