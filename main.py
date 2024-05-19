import openpyxl
import polib

def build_translation_dict(excel_file, po_file):
    # Read Excel file
    workbook = openpyxl.load_workbook(excel_file)

    translation_dict = {}

    print("Building translation replacement associations...")

    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]

        # Get the range of original and translated text columns
        original_texts = sheet['A']
        translated_texts = sheet['C']

        for original_text_cell, translated_text_cell in zip(original_texts, translated_texts):
            original_text = original_text_cell.value
            translated_text = translated_text_cell.value

            # Skip the line if either original or translated text is empty
            if original_text is None or translated_text is None:
                continue

            # Add original and translated text to the dictionary
            translation_dict[original_text] = translated_text

    return translation_dict

def apply_translations(excel_file, po_file):
    # Build a dictionary of original text to translated text
    translation_dict = build_translation_dict(excel_file, po_file)

    # Read .po file
    po = polib.pofile(po_file)

    print("The following IDs have corrected translations:")

    for entry in po:
        # If the translation is found in the dictionary, replace it
        if entry.msgid in translation_dict:
            print(entry.msgid)
            entry.msgstr = translation_dict[entry.msgid]

    # Save the modified .po file
    po.save(po_file)

    # Print completion message
    print("Completed.")

# Example usage
apply_translations('example.xlsx', 'example.po')
