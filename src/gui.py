from datetime import datetime
from pathlib import Path

# import tkinter as tk
from tkinter.messagebox import showinfo, showerror
from tkinter.filedialog import askopenfilename, askdirectory
from tkinter import Tk, StringVar, Label, Button
from os.path import exists
from os import remove

# absolute import
import src.utilis as utilis
from src.workbook import Workbook



class App(Tk):
    '''Class representing GUI'''
    def __init__(self):
        '''Initializes an App instance.'''
        super().__init__()
        self.title('Copy Excel')
        self.geometry('732x160')
        self.file = 'path.csv'
        self.font = ('Arial', 10)
        self.template_path = StringVar()
        self.target_path = StringVar()
        self.source_path = StringVar()
        self.set_paths()
        self.add_widgets()
        self.place_widgets()
    
    def set_paths(self) -> None:
        '''Sets properties template_path and target_path
        according to the csv file with filepaths.
        '''
        if not utilis.file_exists(self.file): return
        (template_path, target_path) = utilis.get_paths(self.file)
        self.template_path.set(template_path)
        self.target_path.set(target_path)
    
    def add_widgets(self) -> None:
        '''Adds widgets to an App as a property.'''
        self.template_label = Label(self, textvariable=self.template_path, 
                                    font=self.font, width = 80, anchor="w")
        self.target_label = Label(self, textvariable=self.target_path, 
                                  font=self.font, width = 80, anchor="w")
        self.source_label = Label(self, textvariable=self.source_path, 
                                  font=self.font, width = 80, anchor="w")
        self.template_btn = Button(self, text="template", 
                                   command=self.select_template)
        self.target_btn = Button(self, text="target folder", 
                                 command=self.select_target)
        self.source_btn = Button(self, text="source", 
                                 command=self.select_source, 
                                 fg='white', bg='blue')
        self.copy_btn = Button(self, text="copy", command=self.copy_excel_data)

    def place_widgets(self) -> None:
        '''Creates layout of the graphical user interface.'''
        self.template_label.grid(row=0, column=1, sticky='e')
        self.target_label.grid(row=1, column=1, sticky='e')
        self.source_label.grid(row=2, column=1, sticky='e')
        self.rowconfigure(3,weight=2)

        self.template_btn.grid(row=0, column=0, 
                               sticky='nesw', 
                               padx=(10, 2), 
                               pady=(5, 2))
        self.target_btn.grid(row=1, column=0, 
                             sticky='nesw', 
                             padx=(10, 2), 
                             pady=(2, 2))
        self.source_btn.grid(row=2, column=0, 
                             sticky='nesw', 
                             padx=(10, 2), 
                             pady=(2, 2))  
        self.copy_btn.grid(row=3, column=0, 
                           sticky='nesw', 
                           columnspan=2, 
                           pady=(2, 5), 
                           padx=(10, 10))


    def select_file(self, is_file: bool) -> None:
        '''Prompts window to select file(True) or directory(False).'''
        try:
            if is_file: filepath = askopenfilename()
            else: filepath = askdirectory()
            if not exists(filepath): 
                raise RuntimeError()
            return filepath
        except:
            showerror(message='Select correct file/directory')
            return None
           
    ###  Methods related to buttons ###
    def select_template(self) -> None:
        '''Sets teplate_path property and updates a csv file with filepaths.'''
        new_filepath = self.select_file(True)
        if new_filepath:
            utilis.replace_file(self.file, new_filepath, 0)
            self.template_path.set(new_filepath)

    def select_target(self) -> None:
        '''Sets target_path property and updates a csv file with filepaths.'''
        new_filepath = self.select_file(False)
        if new_filepath:
            utilis.replace_file(self.file, new_filepath, 1)
            self.target_path.set(new_filepath)
    
    def select_source(self) -> None:
        '''Sets source_path property.'''
        new_filepath = self.select_file(True)
        if new_filepath: 
            self.source_path.set(new_filepath)
    
    def copy_excel_data(self) -> None:
        '''Copies data between excel files and runs macro to update charts.'''
        # copy files to target location
        now = datetime.now().strftime("%Y%m%d-%H%M%S")
        source_path = utilis.copy_file(Path(self.source_path.get()), 
                                       Path(self.target_path.get()), 
                                       now, False)
        template_path = utilis.copy_file(Path(self.template_path.get()), 
                                         Path(self.target_path.get()), 
                                         now, True)

        #copy data between excel files
        template_book = Workbook(template_path)
        source_book = Workbook(source_path)
        template_book.copy_all_sheets(source_book)
        source_book.close()
        remove(source_path)
