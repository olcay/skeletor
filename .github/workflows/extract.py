import os
import shutil
import logging
logging.basicConfig(level=logging.INFO)
from oletools.olevba3 import VBA_Parser

EXCEL_FILE_EXTENSIONS = ('xlsb', 'xls', 'xlsm', 'xla', 'xlt', 'xlam',)
KEEP_NAME = False  # Set this to True if you would like to keep "Attribute VB_Name"
FILEID_SPECIFIER = 'FileId' # This is the file id to keep track of the code in case of file name changes


def parse(workbook_path):
    logging.info('Working on ' + workbook_path)
    
    vba_path = workbook_path + '.vba'
    vba_parser = VBA_Parser(workbook_path)
    vba_modules = vba_parser.extract_all_macros() if vba_parser.detect_vba_macros() else []

    for _, _, filename, content in vba_modules:
        lines = []
        # Split the module content into lines for a better display
        if '\r\n' in content:
            lines = content.split('\r\n')
        else:
            lines = content.split('\n')
        if lines:
            content = []
            # Check the lines for special attributes
            for line in lines:
                if line.startswith('Attribute') and 'VB_' in line:
                    # Ignore Attribute lines except VB_Name if it is needed
                    if 'VB_Name' in line and KEEP_NAME:
                        content.append(line)
                elif line.startswith('\' ' + FILEID_SPECIFIER):
                    # If an id is defined then use it as a folder name
                    vba_path = line.strip().split(' ')[-1] + '.vba'
                    logging.info(FILEID_SPECIFIER + ' is defined. The target directory is changed to ' + vba_path)
                    content.append(line)
                else:
                    content.append(line)
            if content and content[-1] == '':
                content.pop(len(content)-1)
                non_empty_lines_of_code = len([c for c in content if c])
                if non_empty_lines_of_code > 0:
                    if not os.path.exists(os.path.join(vba_path)):
                        # If the folder does not exists, create a new one
                        os.makedirs(vba_path)
                    # Write the content to the module file
                    with open(os.path.join(vba_path, filename), 'w') as f:
                        f.write('\n'.join(content))


if __name__ == '__main__':
    # Walk thorugh all the files in the root directory
    for root, dirs, files in os.walk('.'):
        for f in dirs:
            # Remove the previous extraction folders
            if f.endswith('.vba'):
                shutil.rmtree(os.path.join(root, f))

        for f in files:
            # Run parse method only for the files with Excel extensions
            if f.endswith(EXCEL_FILE_EXTENSIONS):
                parse(os.path.join(root, f))