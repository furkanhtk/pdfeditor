import tkinter as tk
import tkinter.messagebox
import tkinter.filedialog
import PyPDF2
import os
import win32com.client

root = tk.Tk()

text = tk.Label(root, text="Welcome to PDF Editor")
text.pack(side=tk.TOP)

statusbar = tk.Label(root, text="Furkan Aslan HATIK", relief=tk.GROOVE)  # SUNKEN
statusbar.pack(side=tk.BOTTOM, fill=tk.X)

list_files = []
path_files = []


# Functions

def about_me():
    tk.messagebox.showinfo("Contact", "Please contact via furkanhtk@gmail.com")


def open_browser():
    file_path = tk.filedialog.askopenfilename(
        filetypes=(("Powerpoint files", "*.pptx"), ("Powerpoint files", "*.ppt"), ("all files", "*.*")))
    list_files.append(file_path)
    if list_files[-1] == '':
        list_files.pop(-1)
        information['text'] = " "
    else:
        information['text'] = "{0} added".format(os.path.basename(list_files[-1]))
        add_to_listbox()


def add_to_listbox():
    files_list_box1.insert(tk.END, os.path.basename(list_files[-1]))


def list_box_up():
    selected_file = files_list_box1.curselection()
    selected_file_index = int(selected_file[0])
    if not selected_file:
        return
    for selected in selected_file:
        # skip if item is at the top
        if selected == 0:
            continue
        temp = list_files[selected_file_index]
        list_files.pop(selected_file_index)
        list_files.insert((selected_file_index - 1), temp)
        temp2 = files_list_box1.get(selected)
        files_list_box1.delete(selected)
        files_list_box1.insert(selected - 1, temp2)


def list_box_down():
    selected_file = files_list_box1.curselection()
    selected_file_index = int(selected_file[0])
    if not selected_file:
        return
    for selected in selected_file:
        # skip if item is at the top
        temp = list_files[selected_file_index]
        list_files.pop(selected_file_index)
        list_files.insert((selected_file_index + 1), temp)
        temp2 = files_list_box1.get(selected)
        files_list_box1.delete(selected)
        files_list_box1.insert(selected + 1, temp2)


def delete_from_listbox():
    selected_file = files_list_box1.curselection()
    selected_file_index = int(selected_file[0])
    list_files.pop(selected_file_index)
    files_list_box1.delete(selected_file)


def delete_all_listbox():
    list_files.clear()
    path_files.clear()
    files_list_box1.delete(0, tk.END)


def covert_files():
    temp = []
    powerpoint = win32com.client.Dispatch("Powerpoint.Application")
    if varCheckButton.get() == True:
        save_file_path = tk.filedialog.asksaveasfilename(filetypes=(("Pdf", "*.pdf"), ("all files", "*.*")))
    for x in range(len(list_files)):
        path_files.append(list_files[x].replace("/", "\\"))
        temp.append(path_files[x].replace(".pptx", ".pdf"))
        pdf = powerpoint.Presentations.Open(os.path.abspath(path_files[x]), WithWindow=False)
        pdf.SaveAs(os.path.abspath(temp[x]), 32)
        pdf.Close()
    powerpoint.Quit()
    if varCheckButton.get() == True and varCheckButton2.get() == True:
        merger = PyPDF2.PdfFileMerger()
        for pdf in temp:
            merger.append(pdf)
        merger.write((save_file_path+".pdf"))
        PyPDF2.PdfFileMerger.close(merger)
        my_file = open((save_file_path+".pdf"), 'rb')
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
        out_stream = open((save_file_path+"_blank"+ ".pdf"), 'wb')
        out_pdf.write(out_stream)
        out_stream.close()
        my_file.close()
        temp.clear()

    elif varCheckButton.get() == True and varCheckButton2.get() == False:
        merger = PyPDF2.PdfFileMerger()
        for pdf in temp:
            merger.append(pdf)
        merger.write((save_file_path+".pdf"))
        PyPDF2.PdfFileMerger.close(merger)
    elif varCheckButton.get() == False and varCheckButton2.get() == True:
        for file in temp:
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
            value = file.find(".")
            file = file[:value] + "_blank" + file[value:]
            out_stream = open(file, 'wb')
            out_pdf.write(out_stream)
            out_stream.close()
            my_file.close()
    else:
        pass

    if varCheckButton3.get() == True:
        for file in temp:
            os.remove(file)
        os.remove((save_file_path+".pdf"))
    else:
        pass



# Layout

left_frame = tk.Frame(root)
left_frame.pack(side=tk.LEFT, padx=30)

right_frame = tk.Frame(root)
right_frame.pack(side=tk.RIGHT)

# Create the menubar

menubar = tk.Menu(root)
root.config(menu=menubar)

# Create the submenu

subMenu = tk.Menu(menubar, tearoff=0)
menubar.add_cascade(label="File", menu=subMenu)
subMenu.add_command(label="Exit", command=root.destroy)
subMenu = tk.Menu(menubar, tearoff=0)
menubar.add_cascade(label="Help", menu=subMenu)
subMenu.add_command(label="About me", command=about_me)

# Configure size and icons

# root.geometry('800x500')  # set to size of program
root.title("Pdf Converter")  # title of the program
root.iconbitmap(r'./icons/pdf.ico')  # change icon of program

photo1 = tk.PhotoImage(file="./icons/pdf.png")
add_png = tk.PhotoImage(file="./icons/add.png")
merge_png = tk.PhotoImage(file="./icons/merge.png")
delete_png = tk.PhotoImage(file="./icons/delete.png")
up_png = tk.PhotoImage(file="./icons/up.png")
down_png = tk.PhotoImage(file="./icons/down.png")

# Labels (RIGHT)

button_convert = tk.Button(right_frame, text="CONVERT", command=covert_files)
button_convert.pack()
varCheckButton = tkinter.IntVar()
merge_button = tk.Checkbutton(right_frame, text="Merge Files", variable=varCheckButton)
merge_button.pack()

varCheckButton2 = tkinter.IntVar()
add_blank_page_button = tk.Checkbutton(right_frame, text="Add Blank Pages", variable=varCheckButton2)
add_blank_page_button.pack()

varCheckButton3 = tkinter.IntVar()
add_blank_page_button = tk.Checkbutton(right_frame, text="Don't Save Temp Files ", variable=varCheckButton3)
add_blank_page_button.pack()
# -----------------------------------------------------------------------------------------------------------------------------
if not list_files:
    information = tk.Label(right_frame, text="")
    information.pack(side=tk.BOTTOM, fill=tk.X)

    # List Box (LEFT)
    files_list_box1 = tk.Listbox(left_frame, selectmode=tk.EXTENDED)
    files_list_box1.pack()

# Buttons (LEFT)

button_add = tk.Button(left_frame, image=add_png, command=open_browser)
button_add.pack(side=tk.LEFT)
button_delete = tk.Button(left_frame, image=delete_png, command=delete_from_listbox)
button_delete.pack(side=tk.LEFT)
button_delete2 = tk.Button(left_frame, text="Delete All", command=delete_all_listbox)
button_delete2.pack(side=tk.LEFT)
button_up = tk.Button(left_frame, image=up_png, command=list_box_up)
button_up.pack(side=tk.LEFT)
button_down = tk.Button(left_frame, image=down_png, command=list_box_down)
button_down.pack(side=tk.LEFT)

root.mainloop()
