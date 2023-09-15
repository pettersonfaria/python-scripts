import zipfile
import os
from lxml import etree
import sys

def remove_tag_from_worksheets(input_file, tag_name):
    output_file = input_file.replace('.xlsx', '_nopass.xlsx')
    with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as new_zip:
        with zipfile.ZipFile(input_file, 'r') as zip_ref:
            for item in zip_ref.infolist():
                with zip_ref.open(item.filename, 'r') as original_file:
                    if item.filename.startswith('xl/worksheets/') and item.filename.endswith('.xml'):
                        tree = etree.parse(original_file)
                        for element in tree.iter():
                            if element.tag.endswith(tag_name):
                                element.getparent().remove(element)
                        with new_zip.open(item.filename, 'w') as new_xml_file:
                            tree.write(new_xml_file, xml_declaration=True, encoding='utf-8')
                    else:
                        new_zip.writestr(item.filename, original_file.read())
    return output_file

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print(f"Use: python <remove-pass-from-xlsx.py> <file.xlsx>")
    else:
        input_excel_file = sys.argv[1]
        output_file = remove_tag_from_worksheets(input_excel_file, "sheetProtection")
        print(f"The tag 'sheetProtection' was removed: '{output_file}'.")
