# AutoDocAttach
## Background
The AutoDocAttach project automates the System Functional Acceptance Test (SFAT) by streamlining the attachment process of over 50 devices' artefacts (evidence) across 11 types of test cases. Some test cases necessitate multiple artefacts per device, significantly reducing the tedium of repetitive manual actions. This automation eliminates the need for more than 600 repetitive, time-consuming tasks involving navigating through the 'Insert Object' menu, browsing file paths, setting 'Display as icon,' and confirming with the 'OK' button. The project's aim is to simplify and expedite the attachment of artefacts to every specific test case tables within a MS Word document, enhancing efficiency and productivity in SFAT procedures.

### Description
This project automates the attachment of text files to specified tables in a Word document using Python and the win32com library.

### Usage
1. Install the necessary Python packages:
```
pip install pywin32
```

2. Update the script variables:
- **no_of_table**: Number of unrelavant tables to skip initially.

  For example, these 2 tables came before my test case tables.
  ![image](https://github.com/jl-chan/AutoDocAttach/assets/115695686/125d43d5-80ab-4793-ba02-a8be1fd40fda)

- **parent_directory**: Parent directory containing subdirectories of artefacts.
  ![image](https://github.com/jl-chan/AutoDocAttach/assets/115695686/217f34ec-e222-43ff-84e8-4205978e11fc)

### Configuration
Adjust the script variables `no_of_table` and `parent_directory` to match your specific use case.

### How It Operates
The script utilizes the **win32com** library to interact with Microsoft Word. It reads text files from multiple subdirectories within a parent directory and attaches these files in reverse order to tables in a Word document.

> The script operates as follows:
1. **Word Document Initialization**: The script opens the designated Word document essential for the SFAT procedure.
   
2. **Artefact Collection**: It systematically gathers text files from distinct subdirectories within a parent directory. Each of these subdirectories corresponds to specific test cases.

3. **Table-Specific Attachment**: For each of the 11 distinct test cases, which possess their dedicated tables within the document, the script attaches these artefacts in reverse order. Each test case's subdirectory contains artefacts relevant to that particular test, enhancing the document's organization and aligning artefacts with their corresponding test case.

This automated process eliminates the need for manual intervention, significantly reducing the repetitive and time-intensive task of individually inserting, browsing, and arranging artefacts within the document, thereby streamlining the SFAT documentation workflow. 

### Sample result
![image](https://github.com/jl-chan/AutoDocAttach/assets/115695686/e5686bd3-bd4f-480e-827e-d213c28cbadf)

### To Imporve
1. Able to append new object(s) at the bottom of the cell.

The decision to 'attach artefacts in reverse order' stemmed from the difficulty encountered while appending new objects within the designated, same cell at (3, 4) in the table. The code consistently added objects at the beginning of the cell, disrupting the intended sequence. Rather than appending objects at the bottom of the same cell, a workaround was implemented to preserve the sequence identical to that within the directories. This workaround involves inserting object(s) in a reversed order within the specified cell, ensuring that artefacts align in the desired sequence within the cell (3, 4).

