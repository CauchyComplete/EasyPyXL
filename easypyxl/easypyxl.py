"""
26 Sep 2021
Myung-Joon Kwon (CauchyComplete)
corundum240@gmail.com
"""
import openpyxl
import os
import collections.abc
import time


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
        self.prev_saved_time = time.time()
        self.save_excel()

    class Cursor:
        def __init__(self, workbook_class, sheet, start_cell, seq_len, move_vertical, reader, auto_save, auto_save_time):
            self.workbook_class = workbook_class
            self.sheet = sheet
            self.start_cell = start_cell
            self.move_vertical = move_vertical
            self.item_count = 0
            self.seq_len = seq_len
            self.reader = reader
            self.auto_save = auto_save
            self.auto_save_time = auto_save_time

        def __del__(self):
            self.workbook_class.save_excel()

        def _write_cell(self, val):
            if self.move_vertical:
                self.sheet.cell(self.start_cell[0] + (self.item_count % self.seq_len),
                                self.start_cell[1] + (self.item_count // self.seq_len)).value = val
            else:
                self.sheet.cell(self.start_cell[0] + (self.item_count // self.seq_len),
                                self.start_cell[1] + (self.item_count % self.seq_len)).value = val
            self.item_count += 1

        def write_cell(self, val):
            if self.reader:
                raise PermissionError("EasyPyXL: You cannot write_cell() using a cursor with reader=True")
            if isinstance(val, collections.abc.Sequence) and not isinstance(val, str):
                # list or tuple
                for v in val:
                    self._write_cell(v)
            else:
                # string, int, or float
                self._write_cell(val)

            if self.auto_save:
                self.workbook_class.save_excel(self.auto_save_time)

        def _read_cell(self):
            if self.move_vertical:
                val = self.sheet.cell(self.start_cell[0] + (self.item_count % self.seq_len),
                                      self.start_cell[1] + (self.item_count // self.seq_len)).value
            else:
                val = self.sheet.cell(self.start_cell[0] + (self.item_count // self.seq_len),
                                      self.start_cell[1] + (self.item_count % self.seq_len)).value
            self.item_count += 1
            return val

        def read_cell(self, amount=1):
            if not self.reader:
                raise PermissionError("EasyPyXL: You cannot read_cell() using a cursor with reader=False")
            if amount >= 2:
                result = [self._read_cell() for _ in range(amount)]
            elif amount <= 0:
                raise ValueError("EasyPyXL: Cannot read_cell() - amount must be a positive integer.")
            else:
                result = self._read_cell()
            return result

        def read_line(self, amount=1):
            if not self.reader:
                raise PermissionError("EasyPyXL: You cannot read_line() using a cursor with reader=False")
            if amount >= 2:
                result = [self.read_cell(self.seq_len) for _ in range(amount)]
            elif amount <= 0:
                raise ValueError("EasyPyXL: Cannot read_cell() - amount must be a positive integer.")
            else:
                result = self.read_cell(self.seq_len)
            return result

        def skip_cell(self, amount=1):
            self.item_count += amount

        def skip_line(self, amount=1):
            self.item_count += amount * self.seq_len

    def new_cursor(self, sheetname, start_cell, seq_len, move_vertical=False, overwrite=False, reader=False, auto_save=True, auto_save_time=0):
        if auto_save is False and auto_save_time != 0:
            raise ValueError("auto_save is False but auto_save_time is given.")
        if isinstance(start_cell, str):
            start_cell = openpyxl.utils.cell.coordinate_to_tuple(start_cell)
        if self.empty_file:
            sheet = self.workbook.active
            sheet.title = sheetname
            self.empty_file = False
            if self.verbose:
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
        if prev_cell_value is not None and not overwrite and not reader:
            raise ValueError(f"EasyPyXL: start_cell {start_cell} of '{sheetname}' is not empty! "
                             f"Current value: {str(sheet.cell(*start_cell).value)}. To overwrite, set overwrite=True.")
        cursor = self.Cursor(self, sheet, start_cell, seq_len, move_vertical, reader, auto_save, auto_save_time)
        return cursor

    def save_excel(self, auto_save_time=0):
        if auto_save_time == 0:
            self._save_excel()
        elif time.time() - self.prev_saved_time > auto_save_time:
            self._save_excel()

    def _save_excel(self):
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
        self.prev_saved_time = time.time()
