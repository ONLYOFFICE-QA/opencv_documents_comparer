# compare_documents

## Description

Script to open and compare documents in MS Office

## Requirements:

* Python
* MS Office 2019

## Installing Python Libraries:

pip install -r requirements.txt

## Getting Started

* Turn on the dark theme in ms office


* Change config.py file.
  - Set environment variables
    - version - the document server version (Example: "6.4.1.44")
    - source_doc_folder - path to folders with documents before conversion, the folder names must match the file
      extension(if the file extension in the folder is ".doc", then the folder name must be "doc")
    - converted_doc_folder - path to folders with documents after conversion. The names of folders should have the
      form "```<version>_<extension before conversion>_<extension after conversion>```" (Example: "6.4.1.44_doc_docx").
  - list_of_file_names - array for selective comparison by filename

### Commands for starting a document comparison

#### Compare doc to docx:

```invoke doc-docx``` - To compare document statistics by filenames from "converted_doc_folder", and then compare images
by files with different statistics

```invoke doc-docx --st``` - To compare document statistics by file names from “converted_doc_folder”

```invoke doc-docx --im``` - To compare images by file names from “converted_doc_folder”

```invoke doc-docx --ls``` - To compare images by file names from the array in "config.py"

```invoke doc-docx --df``` - To compare images by files with different statistics

#### Compare ppt to pptx:

```invoke ppt-pptx``` - To compare images by file names from “converted_doc_folder”

```invoke ppt-pptx --ls``` - To compare images by file names from the array in "config.py"

#### Compare xls to xlsx:

```invoke xls-xlsx``` - To compare document statistics by filenames from "converted_doc_folder", and then compare images
by files with different statistics

```invoke xls-xlsx --st``` - To compare document statistics by file names from “converted_doc_folder”

```invoke xls-xlsx --im``` - To compare images by file names from “converted_doc_folder”

```invoke xls-xlsx --ls``` - To compare images by file names from the array in "config.py"
