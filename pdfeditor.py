import PyPDF2

my_file=open('S1.pdf','rb')
pdf_reader = PyPDF2.PdfFileReader(my_file)

print(pdf_reader.numPages)

numPages=pdf_reader.getNumPages()
outPdf=PyPDF2.PdfFileWriter()
outPdf.appendPagesFromReader(pdf_reader)
x = 0
while x < numPages-1:
    y = 2*x+1
    x+=1
    #print(y)
    outPdf.insertBlankPage(index = y)
outPdf.addBlankPage()
outStream=open('layout2.pdf','wb')

outPdf.setPageLayout('/TwoPageLeft')
outPdf.write(outStream)
outStream.close()
my_file.close()
