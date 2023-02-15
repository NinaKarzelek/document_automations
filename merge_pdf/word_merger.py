from pathlib import Path
from PyPDF2 import PdfMerger, PdfReader  # pip install PyPDF2

# Define input directory for the pdf files
pdf_dir = Path(__file__).parent / "pdf_files"

# Define & create output directory
pdf_output_dir = Path(__file__).parent / "outputs"
pdf_output_dir.mkdir(parents=True, exist_ok=True)

pdf_files = list(pdf_dir.glob("*.pdf"))

keys = set([file.name[:30] for file in pdf_files])

# Determine the file name length of the base file
# Example of the base files:
# '902 17.03.2022 2000004496.pdf', '904 17.03.2022 2000004497.pdf'
BASE_FILE_NAME_LENGTH = 20

# Define the desired order of the pdf files with specific key
pdf_order = {'FR121-ARP-DC-XX-SP-M-HVAC-8005_Coverpage_Word P02': ['FR121-ARP-DC-XX-SP-M-HVAC-8005--Fans-P02'],
             'FR121-ARP-DC-XX-SP-M-HVAC-8004_Coverpage_Word P02': [
                 'FR121-ARP-DC-XX-SP-M-HVAC-8004--Air Handling Unit-P02']}

for key in keys:
    merger = PdfMerger()
    for file in pdf_files:
        if file.name.startswith(key):
            merger.append(PdfReader(str(file), "rb"))
            if len(file.name) >= BASE_FILE_NAME_LENGTH:
                base_file_name = file.name
    merger.write(str(pdf_output_dir / base_file_name))
    merger.close()
