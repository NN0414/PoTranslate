import openpyxl
import polib

def build_translation_dict(excel_file):
    # Read the Excel file
    workbook = openpyxl.load_workbook(excel_file)

    translation_dict = {}

    print("Building translation replacement associations...")

    # Iterate over all worksheets and collect original and translated text pairs
    all_texts = (
        (original_text_cell.value, translated_text_cell.value)
        for sheet_name in workbook.sheetnames
        for original_text_cell, translated_text_cell in zip(
            workbook[sheet_name]['A'], workbook[sheet_name]['C']
        )
        if original_text_cell.value is not None and translated_text_cell.value is not None
    )

    # Build translation dictionary directly from the worksheet data
    translation_dict.update(all_texts)

    return translation_dict

def apply_translations(excel_file, po_file):
    # Build dictionary of original text to translated text
    translation_dict = build_translation_dict(excel_file)

    # Convert translation_dict to a set for faster lookups
    translation_set = set(translation_dict.keys())

    # Read the .po file
    po = polib.pofile(po_file)

    print("The following IDs have corrected translations:")

    for entry in po:
        # If a corresponding translation is found in the dictionary, replace the translation
        if entry.msgid in translation_set:
            print(entry.msgid)

            if entry.msgid_plural:
                # For msgid_plural, replace all msgstrs
                for idx, msgstr in enumerate(entry.msgstr_plural.values()):
                    entry.msgstr_plural[idx] = translation_dict[entry.msgid]
            else:
                entry.msgstr = translation_dict[entry.msgid]

            # Remove the processed entry from the set
            translation_set.remove(entry.msgid)

        # Check if the translation set is empty
        if not translation_set:
            break

    # Save the modified .po file
    po.save(po_file)

    # Print completion message
    print("Completed")

# Example usage
apply_translations('example.xlsx', 'global.po')
