from tqdm import tqdm

from libs.word import WordCompareImg, Word
from var import *

if __name__ == "__main__":
    for i in tqdm(range(1)):
        # CompareImage()
        # Word(os.listdir(custom_path_to_document_to))
        WordCompareImg()
        # sb.call(["taskkill", "/IM", "WINWORD.EXE" "/T"])
        # Word.delete(path_to_temp_in_test)
