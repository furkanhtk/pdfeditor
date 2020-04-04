import PyPDF2

my_file = open('S1.pdf', 'rb')
pdf_reader = PyPDF2.PdfFileReader(my_file)
print(pdf_reader.numPages)
numPages = pdf_reader.getNumPages()
#pdf_list = sys.argv[1:] # From input (without spaces!)
pdf_list = ['S1.pdf', 'S2.pdf']

def add_blank_page():
    out_pdf = PyPDF2.PdfFileWriter()
    out_pdf.appendPagesFromReader(pdf_reader)
    x = 0
    while x < numPages - 1:
        y = 2 * x + 1
        x += 1
        out_pdf.insertBlankPage(index=y)
    out_pdf.addBlankPage()
    out_stream = open('output.pdf', 'wb')
    out_pdf.write(out_stream)
    out_stream.close()

def pdf_merge():
    merger = PyPDF2.PdfFileMerger()
    for pdf in pdf_list:
        merger.append(pdf)
        merger.write("merger.pdf")




add_blank_page()
my_file.close()
