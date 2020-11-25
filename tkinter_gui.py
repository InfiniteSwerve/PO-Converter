from tkinter import *
from tkinter.filedialog import askopenfilename, askdirectory
from file_parser import convert


class Converter(Tk):
    def __init__(self):
        super(Converter, self).__init__()
        self.title("PO Converter")
        self.minsize(640, 400)
        self.browse_frame = LabelFrame(self, text="")
        self.browse_frame.grid(column=0, row=1, padx=20, pady=20)

        self.convert_frame = LabelFrame(self, text="")
        self.convert_frame.grid(column=0, row=6, padx=20, pady=20)

        self.export_frame = LabelFrame(self, text="")
        self.export_frame.grid(column=0, row=4, padx=30, pady=30)



        self.browse_button()
        self.convert_button()
        self.default_button()
        self.export_button()

    def export_button(self):
        self.export_button = Button(self.export_frame, text="Select a folder to export to", command=self.folder_dialog)
        self.export_button.grid(column=0, row=4)

    def browse_button(self):
        self.browse_button = Button(self.browse_frame, text="Browse For A PO", command=self.file_dialog)
        self.browse_button.grid(column=0, row=1)

    def default_button(self):
        default_var = IntVar()
        self.default_button = Checkbutton(self.convert_frame, text="Save for next time", variable=default_var)
        self.default_button.grid(column=1, row=4)

    def convert_button(self):
        self.convert_button = Button(self.convert_frame, text="Convert", command=self.convert)
        self.convert_button.grid(column=0, row=6)

    def folder_dialog(self):
        self.folder = askdirectory(initialdir="/", title="Browse for folder to save to")
        self.folder_label = Label(self.export_frame, text="")
        self.folder_label.grid(column=0, row=5)
        self.folder_label.configure(text="Saving to " + self.folder)

    def file_dialog(self):
        self.filename = askopenfilename(initialdir="/", title="Select A File",
                                        filetype=(("excel files", "*.xls*"), ("all files", "*.*")))
        self.file_label = Label(self.browse_frame, text="")
        self.file_label.grid(column=1, row=2)
        self.file_label.configure(text="Selected " + re.search(r"[^\/]+\.(xls|xlsx)$", self.filename).group())

    def convert(self):
        self.convert_label = Label(self.convert_frame, text="")
        self.convert_label.grid(column=1, row=7)
        try:
            self.convert_label.configure(text=f"Converted to {self.folder}\\Order_Import.xlsx")
            return convert(self.filename)
        except AttributeError:
            self.convert_label.configure(text="Select a PO and export folder")
            pass


root = Converter()
root.mainloop()
