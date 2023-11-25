import os
import win32com.client
from utils import reverse_list

word = win32com.client.Dispatch("Word.Application")
doc = word.Documents.Open(r"C:\Users\xxxx\Downloads\SFAT_Auto\Test_with_Multiple_Tables.docx")

#### TO CHANGE ######
# No of Table(s) to be skipped #
no_of_table = 2
parent_directory = r"C:\Users\xxxx\Downloads\SFAT_Auto"
#### TO CHANGE ######

use_case_directory_paths = [os.path.join(parent_directory,name) for name in os.listdir(parent_directory) if os.path.isdir(os.path.join(parent_directory, name))]

# For all use case directory_path e.g. Test_Case_4.1.1, Test_Case_4.1.2, Test_Case_4.1.3 ...
for i, directory_path in enumerate(use_case_directory_paths):
    text_files_to_attach = [os.path.join(directory_path, file) for file in os.listdir(directory_path) if os.path.isfile(os.path.join(directory_path, file))]
    text_files_to_attach_rev = reverse_list(text_files_to_attach)
    # Access the first use case table in the document # doc.Tables
    table = doc.tables[no_of_table+i]
    for text_file in text_files_to_attach_rev:
        # Insert text file as an object in the first table
        cell = table.Cell(3, 4)  # Assuming you want to insert the object in the third row of the fourth column
        cell.Range.InlineShapes.AddOLEObject(
            ClassType="Package",
            FileName=os.path.abspath(text_file),
            DisplayAsIcon=True
        )

doc.Save()
word.Quit()
