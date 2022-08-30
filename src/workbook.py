from __future__ import annotations
from xlwings import Book
from win32api import MessageBox
from win32con import MB_ICONINFORMATION


class Workbook(Book):
    def __init__(self, file_path: str) -> None:
        '''Initializes a Workbook instance.'''
        super().__init__(file_path)
    
    def get_last_cell(self, sheet_name: str) -> tuple[int, int]:
        '''Returns location of a last cell in the sheet.'''
        XlDirectionRight = 2
        sheet = self.sheets[sheet_name]
        last_col = sheet.range("A:A").end(XlDirectionRight).column
        last_row = (sheet.range('A' + str(sheet.cells.last_cell.row))
            .end('up').row)
        return (last_row, last_col)
        
    def summary_msg_box(
            self, source_sheet_names: list[str], 
            copied_sheet_names: list[str], 
            deleted_sheet_names: list[str]) -> None:
        '''Pompts message about not copied sheets from 
        source and about sheets deleted from template.
        '''
        not_copied_sheet_names = []
        for sheet_name in source_sheet_names:
            if not sheet_name in copied_sheet_names: 
                not_copied_sheet_names.append(sheet_name)
        if not_copied_sheet_names:
            message = ('List of not copied source sheets: ' 
                        + ', '.join(map(str, not_copied_sheet_names)))
            MessageBox(self.app.hwnd, message, 'Info', MB_ICONINFORMATION)
        if deleted_sheet_names:
            message =  ('List of deleted sheets from template: ' 
                        + ', '.join(map(str, deleted_sheet_names)))
            MessageBox(self.app.hwnd, message, 'Info', MB_ICONINFORMATION)

    def copy_all_sheets(self, source_book: Workbook) -> None:
        '''Copies cells from source to the template, 
        if no corresponding sheet in source file:
        deletes sheet from the template file
        '''
        source_sheet_names = [sheet.name for sheet in source_book.sheets]
        copied_sheet_names = []
        deleted_sheet_names = []
        for template_sheet in self.sheets:
            sheet_name = template_sheet.name
            if sheet_name == "Schedule demand": continue
            if not sheet_name in source_sheet_names:
                self.sheets[sheet_name].delete()
                deleted_sheet_names.append(sheet_name)
            else:
                copied_sheet_names.append(sheet_name)
                source_last_cell = source_book.get_last_cell(sheet_name)
                template_last_cell = self.get_last_cell(sheet_name)
                to_clear_last_cell = (template_last_cell[0],
                                      source_last_cell[1])

                # clear template, copy source, paste to template
                template_sheet.range((1,1),to_clear_last_cell).clear()
                source_book.sheets[sheet_name].range((1,1),source_last_cell).copy()
                template_sheet.range((1,1),source_last_cell).paste()

                # run VBA macro inside template file
                self.activate()
                macro = self.macro("update_chart")      
                macro(source_last_cell[0], sheet_name)  

        self.summary_msg_box(
            source_sheet_names, copied_sheet_names, 
            deleted_sheet_names)
        self.save()
        return
