# import the required libraries
import os
from comtypes.client import CreateObject
from PyPDF2 import PdfMerger

# Root directory
rootDir = "C:\\Users\\Nina.Karzelek\\Desktop\\wordtopdf\\words"

# List to store absolute filepaths
docx_files = []

# Traverse the directory tree
for dirName, subdirList, fileList in os.walk(rootDir):
    # Find all docx files
    for fname in fileList:
        if fname.endswith('.docx'):
            # Combine the directory and file names to get the absolute path
            docx_files.append(os.path.join(dirName, fname))

# create the comtypes instance
word = CreateObject('Word.Application')

# create a list for all the PDFs
pdf_list = []

for word_file in docx_files:
    # get the PDF file name
    pdf_file = os.path.splitext(word_file)[0] + '.pdf'

    # add the PDF file name to the list
    pdf_list.append(pdf_file)

    # open the Word file
    doc = word.Documents.Open(word_file)

    # save the PDF
    doc.SaveAs(pdf_file, FileFormat=17)

    # close the Word file
    doc.Close()

# create the single PDF
word.Application.Quit(SaveChanges=0)

# Merge all the PDFs into one
merger = PdfMerger()

for pdf in pdf_list:
    merger.append(pdf)

# save the single PDF
merger.write('./' + "merged_file.pdf")

# close the merger
merger.close()
