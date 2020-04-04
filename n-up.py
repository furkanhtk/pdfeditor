import PyPDF2

pdf_list = ['S1.pdf', 'S2.pdf']


def add_blank_page():
    for file in pdf_list:
        my_file = open(file, 'rb')
        pdf_reader = PyPDF2.PdfFileReader(my_file)
        numPages = pdf_reader.getNumPages()
        out_pdf = PyPDF2.PdfFileWriter()
        out_pdf.appendPagesFromReader(pdf_reader)
        x = 0
        while x < numPages - 1:
            y = 2 * x + 1
            x += 1
            out_pdf.insertBlankPage(index=y)
        out_pdf.addBlankPage()
        out_stream = open(("e"+file), 'wb')
        out_pdf.write(out_stream)
        out_stream.close()
        my_file.close()

def pdf_merge():
    merger = PyPDF2.PdfFileMerger()
    for pdf in pdf_list:
        merger.append(pdf)
        merger.write("merger.pdf")




add_blank_page()

