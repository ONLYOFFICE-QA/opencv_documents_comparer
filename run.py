from tqdm import tqdm

from comapare_doc_docx import *
from var import *

if __name__ == "__main__":
    for i in tqdm(range(1)):
        create_project_dirs()
        # open_document_and_compare(list_file_names_doc_from_compare, extension_from, extension_to)

        open_document_and_compare(os.listdir(custom_path_to_document_to),
                                  extension_from,
                                  extension_to)

        sb.call(["taskkill", "/IM", "WINWORD.EXE" "/T"])
        operation.delete(path_to_temp_in_test)

        # operation.rename_files(custom_path_to_document_to)
        # operation.rename_files(custom_path_to_document_from)
