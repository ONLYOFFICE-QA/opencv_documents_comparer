from tqdm import tqdm

from Compare_Image import CompareImage
from word import Word
from word_compaire_image import WordCompareImg

if __name__ == "__main__":
    for i in tqdm(range(1)):
        CompareImage()
        # Word()
        # WordCompareImg()
        # sb.call(["taskkill", "/IM", "WINWORD.EXE" "/T"])
        # Word.delete(path_to_temp_in_test)

        # operation.rename_files(custom_path_to_document_to)
        # operation.rename_files(custom_path_to_document_from)
