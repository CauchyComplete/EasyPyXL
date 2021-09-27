"""
26 Sep 2021
Myung-Joon Kwon (CauchyComplete)
corundum240@gmail.com
"""
import openpyxl
import os
import collections.abc


class Workbook:
    def __init__(self, excel_filepath, verbose=True, backup=True):
        if '.' not in excel_filepath:
            excel_filepath = excel_filepath + ".xlsx"
        else:
            if os.path.splitext(excel_filepath)[-1] != ".xlsx":
                raise IOError(f"EasyPyXL only supports '.xlsx' file. Given: {excel_filepath}")
        if os.path.isfile(excel_filepath):
            workbook = openpyxl.load_workbook(excel_filepath, read_only=False)
            if verbose:
                print(f"EasyPyXL loaded workbook: {excel_filepath}")
            self.empty_file = False
            if backup:
                filename = os.path.splitext(excel_filepath)[0] + "_easypyxl_backup"
                for i in range(0, 10000000):
                    new_filepath = f"{filename}_{i}.xlsx"
                    if not os.path.isfile(new_filepath):
                        break
                else:
                    raise IOError(f"EasyPyXL cannot backup: {excel_filepath}")
                if verbose:
                    print(f"EasyPyXL backup file: {new_filepath}")
                workbook.save(new_filepath)
        else:
            workbook = openpyxl.Workbook()
            if verbose:
                print(f"EasyPyXL created workbook: {excel_filepath}")
            self.empty_file = True
            backup = False

        self.workbook = workbook
        self.excel_filepath = excel_filepath
        self.verbose = verbose
        self.saved_error_counter = 0
        self.backup = backup
        self.save_excel()

    class Cursor:
        def __init__(self, workbook_class, sheet, start_cell, seq_len, move_vertical):
            self.workbook_class = workbook_class
            self.sheet = sheet
            self.start_cell = start_cell
            self.move_vertical = move_vertical
            self.item_count = 0
            self.seq_len = seq_len

        def _write_cell(self, val):
            if self.move_vertical:
                self.sheet.cell(self.start_cell[0] + (self.item_count % self.seq_len),
                                self.start_cell[1] + (self.item_count // self.seq_len)).value = val
            else:
                self.sheet.cell(self.start_cell[0] + (self.item_count // self.seq_len),
                                self.start_cell[1] + (self.item_count % self.seq_len)).value = val
            self.item_count += 1
            self.workbook_class.save_excel()

        def write_cell(self, val):
            if isinstance(val, collections.abc.Sequence) and not isinstance(val, str):
                # list or tuple
                for v in val:
                    self._write_cell(v)
            else:
                # string, int, or float
                self._write_cell(val)

        def skip_cell(self, amount=1):
            self.item_count += amount

    def new_cursor(self, sheetname, start_cell, seq_len, move_vertical=False, overwrite=False):
        if isinstance(start_cell, str):
            start_cell = openpyxl.utils.cell.coordinate_to_tuple(start_cell)
        if self.empty_file:
            sheet = self.workbook.active
            sheet.title = sheetname
            print(f"EasyPyXL created sheet: {sheetname}")
        else:
            if sheetname in self.workbook.sheetnames:
                sheet = self.workbook[sheetname]
                if self.verbose:
                    print(f"EasyPyXL loaded sheet: {sheetname}")
            else:
                sheet = self.workbook.create_sheet(sheetname)
                if self.verbose:
                    print(f"EasyPyXL created sheet: {sheetname}")
        prev_cell_value = sheet.cell(*start_cell).value
        if prev_cell_value is not None and not overwrite:
            raise ValueError(f"EasyPyXL: start_cell {start_cell} of '{sheetname}' is not empty! "
                             f"Current value: {str(sheet.cell(*start_cell).value)}. To overwrite, set overwrite=True.")
        cursor = self.Cursor(self, sheet, start_cell, seq_len, move_vertical)
        return cursor

    def save_excel(self):
        try:
            self.workbook.save(self.excel_filepath)
        except PermissionError:
            filename = os.path.splitext(self.excel_filepath)[0]
            if self.saved_error_counter > 0:
                filename = '_'.join(filename.split('_')[:-1])
            for i in range(self.saved_error_counter, 10000000):
                new_filepath = f"{filename}_{i}.xlsx"
                if not os.path.isfile(new_filepath):
                    self.saved_error_counter += 1
                    break
            else:
                raise IOError(f"EasyPyXL cannot save file: {self.excel_filepath}")
            self.excel_filepath = new_filepath
            if self.verbose:
                print(f"EasyPyXL saved to new file: {self.excel_filepath}")
            self.workbook.save(self.excel_filepath)
