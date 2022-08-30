# import os
# import csv
# import shutil
# import tkinter as tk
from pathlib import Path
from datetime import datetime
from os.path import exists
from tkinter.messagebox import showerror
from shutil import copy
from csv import reader

def file_exists(filename: str) -> bool:
    """Checks if file exists, if not creates one."""
    if exists(filename): 
        return True
    with open(filename, 'a') as f:
        f.write("\n")
    return False


def get_paths(filename: str) -> tuple[str]:
    """Retrieves filepaths from file,
    if can't returns empty string.
    """
    template_path, target_path = '', ''
    with open(filename, newline='\n') as f:
        f_reader = list(reader(f, delimiter=';'))
        try:
            template_path = f_reader[0][0]
            if not exists(template_path): template_path = ''
        except: pass
        try:
            target_path = f_reader[1][0]
            if not exists(target_path): target_path = ''
        except: pass
    return (template_path, target_path)


def replace_file(paths: str, input: str, type:int) -> None:
    """Replaces rows with filepaths in the file with paths."""
    with open(paths, 'r') as f:
        data = f.readlines()

    if len(data) == 1 and type == 1: 
        data.append(input + '\n')
    else: 
        data[type] = input + '\n'
        
    with open(paths, 'w') as f:
        data = f.writelines(data)


def copy_file(source_file: Path, target_folder: Path,
              timestamp: datetime, is_template: bool)-> str:
    '''Copies file to target destination, returns path to copied file.'''
    tag = 'Results-' if is_template else 'tmp-'
    filename = tag + timestamp + "." + (str(source_file).split('.')[-1])
    target_path = target_folder / filename

    try:
        copy(source_file, target_path)
    except:
        showerror(message="Error occurred while copying file.")
    return target_path