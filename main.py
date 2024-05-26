import openpyxl
import polib

def build_translation_dict(excel_file):
    # Read the Excel file
    workbook = openpyxl.load_workbook(excel_file)

    translation_dict = {}

    print("Building translation replacement associations...")

    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        
        # Get the range of msgid, msgid_plural and msgstr columns
        msgid = sheet['A']
        msgid_plural = sheet['B']
        msgstr = sheet['D']

        for msgid_text_cell, msgid_plural_text_cell, msgstr_text_cell in zip(msgid, msgid_plural, msgstr):
            msgid_text = msgid_text_cell.value
            msgid_plural_text = msgid_plural_text_cell.value
            msgstr_text = msgstr_text_cell.value

            if msgid_plural_text is None:
                msgid_plural_text = ' '
            
            # Skip the line if either original or translated text is empty
            if msgid_text is None or msgstr_text is None:
                continue

            Pokey = msgid_text + ';' + msgid_plural_text
            translation_dict[Pokey] = msgstr_text
    
    return translation_dict

def apply_translations(excel_file, po_file):
    # Build dictionary of original text to translated text
    translation_dict = build_translation_dict(excel_file)

    # Read the .po file
    po = polib.pofile(po_file)

    print("The following IDs have corrected translations:")

    for entry in po:
        PoKey = (entry.msgid, entry.msgid_plural if entry.msgid_plural else ' ')

        # If a corresponding translation is found in the dictionary, replace the translation
        if PoKey in translation_dict:
            print(PoKey)

            if entry.msgid_plural:
                # For msgid_plural, replace all msgstrs
                for idx, msgstr in enumerate(entry.msgstr_plural.values()):
                    entry.msgstr_plural[idx] = translation_dict[PoKey]
            else:
                entry.msgstr = translation_dict[PoKey]

            # Remove the processed entry from the dictionary
            del translation_dict[PoKey]

        # Check if the translation dictionary is empty
        if not translation_dict:
            break

    # Save the modified .po file
    po.save(po_file)

    # Print completion message
    print("Apply translations Complete.")

# Example usage
apply_translations('example.xlsx', 'global.po')
