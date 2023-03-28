from tkinter import END, ttk
from PyPDF2 import PdfMerger
import tkinter as tk
from tkinter import filedialog as fd, messagebox
from docx2pdf import convert
import sys #_MEIPASS
from os import path, remove
from time import sleep


class windowApp():
    def __init__(self) -> None:
        self.__inputNames = []
        self.__lastIndex = 0

        self.__filetypes = ((r'Plik pdf', r'*.pdf'), (r'Plik docx', r'*.docx'), (r'Plik doc', r'*.doc'))
        # okono aplikacji
        self.__window = tk.Tk()
        # wymiar okna
        self.__window.geometry(r'315x293')
        # nazwa okna
        self.__window.title(r'Scalanie plików')
        # ikona
        self.__window.iconbitmap(self.__resourcePath(r'icon.ico'))
        # self.__window.iconphoto(True, tk.PhotoImage(self.__resourcePath(r'icon.ico')))

        self.__inputNamesField = tk.StringVar(self.__window)

        ttk.Button(self.__window, text=r'Otwórz pliki',
                   command=self.setInputNames).place(x=5, y=8)

        self.__listbox = tk.Listbox(self.__window, height=15, width=42,
                                    listvariable=self.__inputNamesField)
        self.__listbox.place(x=5, y=45)

        ttk.Button(self.__window, text=r'Zapisz',
                   command=self.__save).place(x=180, y=8)

        ttk.Button(self.__window, text=r'ᐱ', width=2,
                   command=self.__up).place(x=270, y=45)  # up list element

        ttk.Button(self.__window, text=r'ᐯ', width=2,
                   command=self.__down).place(x=270, y=255)  # down list element

        style = ttk.Style(self.__window)
        self.__window.tk.call(r'source', self.__resourcePath(r'theme/azure dark.tcl'))
        style.theme_use(r'azure')
        style.configure(r'Accentbutton', foreground=r'white')
        style.configure(r'Togglebutton', foreground=r'white')
        self.__window.mainloop()

    def __switchElement(self, index, indexChange):
        self.__inputNames[index], self.__inputNames[indexChange] = self.__inputNames[indexChange], self.__inputNames[index]
        self.__listbox.select_clear(0, END)
        self.__listbox.selection_set(indexChange)

    def __setInputNamesField(self):
        fileName = [path.rsplit(r'/', 1)[1] for path in self.__inputNames]
        self.__inputNamesField.set(fileName)

    def setInputNames(self):
        documents_folder = path.join(path.expanduser(r"~"), r"Documents")
        tmpOldNames = self.__inputNames
        tmpNewNames = fd.askopenfilenames(
            title=r'Otwórz pliki pdf', initialdir=documents_folder, filetypes=self.__filetypes)
        [tmpOldNames.append(x) for x in tmpNewNames]
        self.__inputNames = tmpOldNames
        self.__setInputNamesField()
        self.__lastIndex = len(self.__inputNames) - 1

    def __up(self):
        selectedId = self.__listbox.curselection()
        if len(selectedId) > 0:
            selectedId = selectedId[0]
            if selectedId > 0:
                self.__switchElement(selectedId, selectedId-1)
                self.__setInputNamesField()
            else:
                self.__switchElement(selectedId, self.__lastIndex)
                self.__setInputNamesField()

    def __down(self):
        selectedId = self.__listbox.curselection()
        if len(selectedId) > 0:
            selectedId = selectedId[0]
            if selectedId < self.__lastIndex:
                self.__switchElement(selectedId, selectedId+1)
                self.__setInputNamesField()
            else:
                self.__switchElement(selectedId, 0)
                self.__setInputNamesField()

    def __save(self):
        merger = PdfMerger(strict=False)
 
        file_to_remove = []
        folder_selected = tk.filedialog.asksaveasfilename(defaultextension=r"*.pdf", filetypes=((r'Plik pdf', r'*.pdf'),))
        for file in self.__inputNames:
            if file.endswith(r'.docx') or file.endswith(r'.doc'):
                convert(file)
                file = file.replace(path.splitext(file)[1], r'.pdf')
                file_to_remove.append(file)
            merger.append(file)
        merger.write(folder_selected)
        merger.close()
        messagebox.showinfo("Info", "Scalono wszystkie pliki")
        sleep(0.5)
        try:
            [remove(file) for file in file_to_remove]
        except PermissionError:
            messagebox.showinfo("Nie usunięto tymczasowych plików", '\n'.join(file_to_remove))

    @staticmethod
    def __resourcePath(relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = path.abspath(r'.')

        return path.join(base_path, relative_path)


windowApp()
