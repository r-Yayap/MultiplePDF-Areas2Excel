# PDF-Text-Extractor
A PDF text extractor built using python to extract text on specific areas for multiple files

This code as mentioned above can read text in a PDF file by specifically selecting an area in a pdf then extract the texts on that area on multiple PDFs. The output would then be saved as Excel files. It works great on multiple PDF with the same PDF dimensions.


I am not a full-pledged coder so the code might look messy in the eyes of a coder. This is a simple tool to help me and my team in our daily tasks of scanning through PDFs. Regardless, the current capability of this code would suffice for what I had in mind when this was still an idea. It has helped me saved a lot of time and be more efficient, and I hope for those who will use this would also feel the same.

For the meantime, it cannot extract texts in an image due to limited time for coding. This might be done in future development propably by using pytesseract or other OCR modules available. There is a quick guide on how to use it though I hope the UI would be easy to understand to those who would use this.

## Installation
Install the dependencies stated in requirements.txt
```
pip install -r requirements.txt
```

or you could also just make an .exe through pyinstaller

## Usage
1. Run Extract_GUI.py
2. After running the code:
    (1) Browse your desired folder which contains your PDFs.
   ![image](https://github.com/Yayap-dev/PDF-Text-Extractor/assets/21073411/66b54df5-ca84-4367-a3ee-5da2ddba268d)

    (2) Open one sample PDF for us to see where we will extract the texts
   
    (3) Select output location of the Excel File
   
    (4) Draw Rectangles by doing a _Click-drag _motion. This rectangles are the ones that will be extracted.
   
    (5) Press Extract button. **optional: check the 'Include Subfolders?' if you want to include the PDFs in the subfolders

## Notes

- Areas are edittable, double click the table beside the Extract button to edit it.
  
- "Clear Areas" deletes all rendered rectangles/areas

- The "**PDF DWG list**" button lists the PDF and DWG in a directory and checks the filenames of both type to see if a PDF's filename has an exact filename of a DWG. (Output is in excel)

