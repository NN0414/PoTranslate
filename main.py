import openpyxl
import polib

def build_translation_dict(excel_file, po_file):
    # Read the Excel file
    workbook = openpyxl.load_workbook(excel_file)

    translation_dict = {}

    print("Building translation replacement associations...")

    all_texts = []

    # Read all original and translated text data from all worksheets at once
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]

        # Get ranges of original and translated columns
        original_texts = sheet['A']
        translated_texts = sheet['C']

        all_texts.extend(zip(original_texts, translated_texts))

    # Build translation dictionary
    translation_dict.update(
        (original_text_cell.value, translated_text_cell.value)
        for original_text_cell, translated_text_cell in all_texts
        if original_text_cell.value is not None and translated_text_cell.value is not None
    )

    return translation_dict

def apply_translations(excel_file, po_file):
    # Build dictionary of original text to translated text
    translation_dict = build_translation_dict(excel_file, po_file)

    # Read the .po file
    po = polib.pofile(po_file)

    print("The following IDs have corrected translations:")

    for entry in po:
        # If a corresponding translation is found in the dictionary, replace the translation
        if entry.msgid in translation_dict:
            print(entry.msgid)

            if entry.msgid_plural:
                # For msgid_plural, replace all msgstrs
                entry.msgstr_plural[0] = translation_dict[entry.msgid]
                entry.msgstr_plural[1] = translation_dict[entry.msgid]
            else:
                entry.msgstr = translation_dict[entry.msgid]

    # Save the modified .po file
    po.save(po_file)

    # Print completion message
    print("Completed")

# Example usage
apply_translations('example.xlsx', 'global.po')
