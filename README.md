# PDF-Text-Extractor
A PDF text extractor built using python to extract text on specific areas for multiple files.

This tool utilizes [PyMuPDF](https://github.com/pymupdf/PyMuPDF) module to read text in a PDF file, allowing users to select specific areas via a Graphical User Interface for extraction. The extracted text is then saved as an Excel file. It performs efficiently, particularly on multiple PDFs with identical dimensions.

Built with the help of ChatGPT, Google, and Youtube, this is a simple tool to help me and my team in our daily tasks of scanning through PDFs. I am not really a coder so the code might look messy in the eyes of a coder. Regardless, the current capability of this code would suffice for what I had in mind when this was still an idea. It has helped me save a lot of time and be more efficient, and I hope for those who will use this would also feel the same.

Currently, the tool does not support text extraction from images due to time constraints. However, future development may include this feature, possibly integrating pytesseract or other OCR modules. There is a quick guide on how to use it though I hope the UI would be easy to understand to those who would use this.

## Installation
Install the dependencies stated in requirements.txt
```
pip install -r requirements.txt
```
or
```
pip install ﻿customtkinter==5.2.2
pip install matplotlib==3.8.2
pip install openpyxl==3.2.0b1
pip install pandas==2.1.4
pip install pillow==10.2.0
pip install PyMuPDF==1.23.11
```
## Usage
1. Run Extract_GUI.py ( or you could also just make an .exe through pyinstaller, then run it )
2. Then:
   
    (1) Browse the desired folder which contains your PDFs.
   
   ![image](https://github.com/Yayap-dev/PDF-Text-Extractor/assets/21073411/66b54df5-ca84-4367-a3ee-5da2ddba268d)

    (2) Open one sample PDF for us to see where we will extract the texts
   
    (3) Select output location of the Excel File
   
    (4) Draw Rectangles by doing a _Click-drag _motion. These rectangles are the ones that will be extracted.
   
    (5) Press Extract button.

## Notes

- Check the 'Include Subfolders?' if you want to include the PDFs in the subfolders
  
- The "**PDF DWG list**" button lists the PDF and DWG in a directory and checks the filenames of both type to see if a PDF's filename has an exact filename of a DWG. (Output is in excel)

- Coordinates are edittable, double click the table beside the Extract button to edit it.
  
- "Clear Areas" deletes all rendered rectangles/areas

## Links

[LinkedIn Profile](https://github.com/pymupdf/PyMuPDF)
