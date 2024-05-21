# Po Translate

## Introduction

This is a Python script designed to assist with translations by applying translations from an Excel file to a gettext PO file.

## Requirements

- Python 3.x
- openpyxl
- polib

## Installation

1. Clone this repository to your local machine:

    ```
    git clone https://github.com/your_username/translation-tool.git
    ```

2. Install the required Python packages:

    ```
    pip install openpyxl polib
    ```

## Usage

### 1. Prepare Excel File

Prepare an Excel file with the following structure:

| Original Text | Context | Translated Text |
|---------------|---------|-----------------|
| Hello         | Greeting| 你好             |
| ...

Save this Excel file for later use.

### 2. Apply Translations

To apply translations from the Excel file to a gettext PO file, run the script `apply_translations.py` with the following command:

```
python apply_translations.py example.xlsx global.po
```

Replace `example.xlsx` with the path to your Excel file and `global.po` with the path to your gettext PO file.

## Functionality

The script provides two main functions:

1. **build_translation_dict**: Reads the Excel file and builds a dictionary of original text to translated text.

2. **apply_translations**: Applies translations from the dictionary to the specified gettext PO file.

## Example

```python
apply_translations('example.xlsx', 'global.po')
```

This will apply translations from `example.xlsx` to `global.po`.
