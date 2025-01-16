# Document testing project

## Description

This project aims to test the conversion,
opening files after conversion, and comparing documents.

## Requirements

* Python 3.12
* MS Office 2019
* LibreOffice Community Version: 7.3.0.3

## Installing Python Libraries

### Install Poetry Package Manager

Instruction: [Python Poetry Installation Guide](https://python-poetry.org/docs/#installation)

## Getting Started

* MS office settings
  * Turn on the dark theme in ms office and in LibreOffice
  * In MS PowerPoint change
    ```Options\Advanced\Display\Open all documents using this views```
    on ```Normal - slide only```

* Change config.py file.
  * Set environment variables
    * version - the document server version (Example: "8.1.0.50")
    * source_doc_folder - path to folders with documents before conversion,
    the folder names must match the file
      extension
      (if the file extension in the folder is ".doc",
      then the folder name must be "doc")
    * converted_doc_folder - path to folders with documents after conversion.
      The names of folders should have the
      form "`<version>_(dir_<extension before conversion>
_<extension after conversion>)_(os_<operation system>)_(mode_<Default or t-format>)`"
      Example: `8.1.0.50_(dir_xps-docx)_(os_linux)_(mode_Default)`.
    * list_of_file_names - array for selective comparison by filename

### Sending messages to Telegram

* To send termination reports of the script to Telegram, you need:
* to add 2 environment variables
  * `TELEGRAM_TOKEN` - with a token from your bot on telegram
  * `CHANNEL_ID` - with the channel id on telegram
* or 2 files in the path `~/.telegram`:
  * `token` - with a token from your bot on telegram
  * `chat` - with the channel id on telegram
* to send messages via proxy, add a file at the path `~/.telegram/proxy.json` containing:

```json
{
  "login": "",
  "password": "",
  "ip": "",
  "port": ""
}
```

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

### Download files

`invoke download-files`

#### Download files flags

`--cores` or `-c` - The amount CPU cores
to use for parallel downloading. (default: None)

### Upload files

`invoke s3-upload`

#### Upload files flags

`--dir-path` or `-d` - The local directory path
containing files to be uploaded.

`--cores` or `-c` - The amount CPU cores
to use for parallel downloading. (default: None)
