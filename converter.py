
import os
import sys
import glob
import win32com.client
import time

list_path =['C:/rojects/pdfeditor/1.pptx', 'C:\Projects\pdfeditor\2.pptx']
path_of_inputfile = r'C:\Projects\pdfeditor\1.pptx'
path_of_outputfile = r'C:\Projects\pdfeditor\deneme.pdf'
path=os.path.abspath(list_path[0])
path_output=os.path.abspath(path_of_outputfile)
powerpoint = win32com.client.Dispatch("Powerpoint.Application")
pdf = powerpoint.Presentations.Open(path, WithWindow=False)
pdf.SaveAs(path_of_outputfile, 32)
pdf.Close()
powerpoint.Quit()





