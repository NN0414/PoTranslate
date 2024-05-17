import openpyxl
import polib

def apply_translations(excel_file, po_file):
    # Read Excel file
    workbook = openpyxl.load_workbook(excel_file)

    # Read .po file
    po = polib.pofile(po_file)

    print("The following IDs have been corrected translations")
    
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]

        # Get the range of original text and translated text columns
        original_texts = sheet['A']
        translated_texts = sheet['C']

        # Process each original text
        for row_idx, original_text_cell in enumerate(original_texts, start=1):
            original_text = original_text_cell.value

            # Skip the row if the original text cell is empty
            if original_text is None:
                continue
                
            # Look up the translation corresponding to this original text in the .po file
            for entry in po:
                if entry.msgid == original_text:
                    print(original_text)
                    
                    # Update the translation
                    entry.msgstr = translated_texts[row_idx - 1].value
                    

    # Save the modified .po file
    po.save(po_file)

    # Print completion message
    print("Completed")

# Example usage
apply_translations('example.xlsx', 'example.po')
