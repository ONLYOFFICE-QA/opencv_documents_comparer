# Document testing project

## Description

A project for testing conversion opening files after conversion and comparing documents.

## Requirements

* Python 3.11
* MS Office 2019
* LibreOffice Community Version: 7.3.0.3

## Installing Python Libraries

### First you need to install the package manager poetry

* Unix system
  * Run command:
  `curl -sSL <https://install.python-poetry.org> | python3 -`
  * Add Poetry to your PATH
  `$HOME/.local/bin`
* Windows
  * Run command in powershell:
  `(Invoke-WebRequest -Uri https://install.python-poetry.org
   -UseBasicParsing).Content | py -`
  * Add Poetry to your PATH
  `%APPDATA%\Python\Scripts`

To install the dependencies via poetry, run the command
`poetry install`

To activate the virtual environment, run the command
`poetry shell`

### If installing dependencies is only possible via pip

* To install the dependencies via pip:
`python3 install_requirements.py`

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

* Sending messages to Telegram
* To send termination reports of the script to Telegram,
  you need to add 2 environment variables
  * TELEGRAM_TOKEN - with a token from your bot on telegram
  * CHANNEL_ID - with the channel id on telegram

### Command for starting a document comparison

`invoke compare-test` - To compare images by file names
from "converted_doc_folder"

#### Compare test flags

`--direction` or `-d` - specifies which files to compare.
Example: `-d doc-docx` to compare doc-docx files.

`--telegram` or `-t` - to send reports in a telegram.

`--ls` or `-l` - to open files from files_array in config.py

### Command for make files for openers test via x2ttester

`invoke make-files`

#### Make files flags

`--telegram` or `-t` - to send reports in a telegram.

`--direction` or `-d` - Specifies the direction for converting files.
Example: `-d doc-docx`

`--version` or `-v` - Specifies the version x2t libs.
Example: `invoke make-files -v 7.4.0.163`

### Command for starting openers

`invoke open-test` - Running file opening tests

#### Open test flags

`--direction` or `-d` - specifies which files to open.
Example: `-d docx` to open only docx files. if no required - all
files will be opened

`--ls` or `-l` - to open files from files_array in config.py

`--path "./path/to/your/folder"` - to open files in the specified folder.

`--telegram` or `-t` - to send reports in a telegram.

`--fast_test` or `-f` - to run tests without taking
into account the tested files

### Command for starting conversion tests

`invoke conversion-test`

#### Conversion test flags

`--direction` or `-d` - Specifies the direction for converting files.
Example: `-d doc-docx`. if not required -
conversion to all formats will be performed.

`--ls` or `-l` - to conversion files from files_array in config.py

`--telegram` or `-t` - to send reports in a telegram.

`--version` or `-v` - Specifies the version x2t libs.

### Command for download x2t libs

`invoke download-core`

#### Download x2t libs flags

`--version` or `-v` - Specifies the version x2t libs.
