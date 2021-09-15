import subprocess as sb

from tqdm import tqdm

from libs.Word_image_compare import WordCompareImg
from var import *

if __name__ == "__main__":
    for i in tqdm(range(1)):
        WordCompareImg(list_file_names_doc_from_compare)
        sb.call(f'powershell.exe kill -Name WINWORD', shell=True)
