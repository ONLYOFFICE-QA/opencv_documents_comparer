# compare_documents

## Description

Script to open and compare documents in MS Office

## Requirements

* Python 3.10
* MS Office 2019
* LibreOffice Community Version: 7.3.0.3

## Installing Python Libraries

pip install -r requirements.txt

## Getting Started

* MS office settings
  * Turn on the dark theme in ms office and in LibreOffice
  * In MS PowerPoint change
    ```Options\Advanced\Display\Open all documents using this views```
    on ```Normal - slide only```

* Change config.py file.
  * Set environment variables
    * version - the document server version (Example: "6.4.1.44")
    * source_doc_folder - path to folders with documents before conversion,
    the folder names must match the file
      extension
      (if the file extension in the folder is ".doc",
      then the folder name must be "doc")
    * converted_doc_folder - path to folders with documents after conversion.
    The names of folders should have the
      form "```<version>_<extension before conversion>_<extension after conversion>```"
      (Example: "6.4.1.44_doc_docx").
  * list_of_file_names - array for selective comparison by filename

### Commands for starting a document comparison

#### Compare doc to docx

```invoke doc-docx``` - To compare images by file names from "converted_doc_folder"

```invoke doc-docx --st``` - To compare document statistics by file names from "converted_doc_folder"

```invoke doc-docx --full``` - To compare document statistics by filenames from
"converted_doc_folder", and then compare images by files with different statistics

```invoke doc-docx --ls``` - To compare images by file names from the array in "config.py"

```invoke doc-docx --df``` - To compare images by files with different statistics

```invoke doc-docx --cl``` - To compare images by file names from the clipboard

#### Compare ppt to pptx

```invoke ppt-pptx``` - To compare images by file names from "converted_doc_folder"

```invoke ppt-pptx --ls``` - To compare images by file names from the array in "config.py"

```invoke ppt-pptx --cl``` - To compare images by file names from the clipboard

#### Compare xls to xlsx

```invoke xls-xlsx``` - To compare images by file names from "converted_doc_folder"

```invoke xls-xlsx --st``` - To compare document statistics by file names from "converted_doc_folder"

```invoke xls-xlsx --full``` - To compare document statistics by filenames from
"converted_doc_folder", and then compare images by files with different statistics

```invoke xls-xlsx --ls``` - To compare images by file names from the array in "config.py"

```invoke xls-xlsx --cl``` - To compare images by file names from the clipboard

#### Compare odp to pptx

```invoke odp-pptx``` - To compare images by file names from "converted_doc_folder"

```invoke odp-pptx --ls``` - To compare images by file names from the array in "config.py"

```invoke odp-pptx --cl``` - To compare images by file names from the clipboard

### Commands for starting openers

`invoke opener-docx` - To open docx with ms Word

`invoke opener-xlsx` - To open xlsx with ms Excel

`invoke opener-pptx` - To open docx with ms PP

`invoke opener-ods` - To open ods with LibreOffice

`invoke opener-odt` - To open odt with LibreOffice

`invoke opener-odp` - To open odp with LibreOffice

To check the opening of a particular conversion direction,
add a flag with the source extension. example:

`invoke opener-docx --doc` - Checking document opening after doc=>docx conversion
