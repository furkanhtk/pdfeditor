import tkinter as tk
from tkinter import *

root = Tk()


root.geometry('800x500') #set to size of program
root.title("Pdf Coverter") # title of the program
root.iconbitmap(r'./icons/pdf.ico') # change icon of program


text = Label(root, text = "Welcome to PDF editor")
text.pack()

photo = PhotoImage(file="./icons/pdf.png")
labelphoto = Label(root , image = photo)
labelphoto.pack()









root.mainloop()

