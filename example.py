import os
from PDFautomatic import BoomPdf

#Class Initialization
pdf = BoomPdf()
fname = f'{os.path.dirname(__file__)}' #Gets working directory.

#Merge Example
#This could be automated to read each file in a directory.
lst = [f'{fname}\\pdf1.pdf', f'{fname}\\pdf2.pdf'] 
pdf.pdf_merger(lst)

#Split example
pdf.pdf_splitter(fname)