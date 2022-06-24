from PyPDF2 import PdfMerger
from PyPDF2 import PdfFileReader
from PyPDF2 import PdfFileWriter

class BoomPdf:
    def __init__(self):
        self.pdfMerger = PdfMerger()
        self.pdfWriter = PdfFileWriter()
    
    def pdf_merger(self, pdf_files):
        """Takes a list of PDF files,
            produces one consolidated PDF file."""
        for file in pdf_files:
            self.pdfMerger.append(file)

        self.pdfMerger.write("output_file.pdf")
        self.pdfMerger.close()
    
    def pdf_splitter(self, pdf_file):
        """Takes a multiple page PDF file,
            splits the file and produces each page
            as individual PDF files."""
        pdf = PdfFileReader(pdf_file)
        for pag in range(pdf.getNumPages()):
            curr_pag = pdf.getPage(pag)
            self.pdfWriter.addPage(curr_pag)
            outputFilename = "Split-Page-{}.pdf".format(pag + 1)
            
            with open(outputFilename, "wb") as out:
                self.pdfWriter.write(out)




