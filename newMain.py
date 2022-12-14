from tkinter import END, ttk
from PyPDF2 import PdfFileMerger
import tkinter as tk
from tkinter import filedialog as fd
import sys #_MEIPASS
from os import path



class windowApp():
    def __init__(self) -> None:
        self.__inputNames = []
        self.__lastIndex = 0

        self.__filetypes = ((r'Plik pdf', r'*.pdf'),)
        # okono aplikacji
        self.__window = tk.Tk()
        # wymiar okna
        self.__window.geometry(r'315x293')
        # nazwa okna
        self.__window.title(r'Scalanie plików PDF')
        # ikona
        self.__window.iconbitmap(self.__resourcePath(r'icon.ico'))

        self.__inputNamesField = tk.StringVar(self.__window)

        ttk.Button(self.__window, text=r'Otwórz pliki pdf',
                   command=self.setInputNames).place(x=5, y=40)

        self.__listbox = tk.Listbox(self.__window, height=13, width=42,
                                    listvariable=self.__inputNamesField)
        self.__listbox.place(x=5, y=75)

        tk.Label(self.__window, text=r'Nazwa pliku wyjściowego:').place(x=5, y=10)
        # output name
        self.__outputName = tk.StringVar(self.__window)
        self.__outputName.set(r'Plik wyjściowy')
        self.__outputName.trace(r'w', self.setOutputName)
        ttk.Entry(self.__window, textvariable=self.__outputName,
                  width=24).place(x=149, y=5)

        ttk.Button(self.__window, text=r'Zapisz',
                   command=self.__save).place(x=180, y=40)

        ttk.Button(self.__window, text=r'ᐱ', width=2,
                   command=self.__up).place(x=270, y=75)  # up list element

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
        tmpOldNames = self.__inputNames
        tmpNewNames = fd.askopenfilenames(
            title=r'Otwórz pliki pdf', initialdir=r'/', filetypes=self.__filetypes)
        [tmpOldNames.append(x) for x in tmpNewNames]
        self.__inputNames = tmpOldNames
        self.__setInputNamesField()
        self.__lastIndex = len(self.__inputNames) - 1

    def setOutputName(self, *args):
        self.__outputName.set(self.__outputName.get())

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
        merger = PdfFileMerger(strict=False)

        [merger.append(pdf) for pdf in self.__inputNames]

        merger.write(f'{self.__outputName.get()}.pdf')

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
