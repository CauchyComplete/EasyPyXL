# EasyPyXL
 This python package is a wrapper of OpenPyXL for easy usage.

You can easily write your data on an Excel(.xlsx) file.

It is especially helpful when used in loops.

Stop wasting your energy by printing your experiment results to a text file and copy-pasting them to an Excel file one by one. Instead, use this package.

I tried to handle possible errors, but if you face one, please report it to Github Issues.

## Install
Install this package:

```pip install easypyxl```

If it does not work, try:

```pip install git+https://github.com/CauchyComplete/EasyPyXL```

## Example 1 : Basics
```angular2html
import easypyxl
workbook = easypyxl.Workbook("my_excel.xlsx")  # Excel file to write on. If the file does not exist, it will be created.
cursor = workbook.new_cursor("MySheet", "A2", 5)  # New cursor at sheet "MySheet", starting from "A2", new line every 5 writes.
for i in range(25):
    cursor.write_cell(i)
```
![ex1](https://github.com/CauchyComplete/EasyPyXL/blob/main/images/ex1.png?raw=true)

## Example 2 : write_cell(list)
```angular2html
import easypyxl
workbook = easypyxl.Workbook("my_excel.xlsx", verbose=False)  # Use verbose=False if you want this package to print only important messages. 
cursor = workbook.new_cursor("MySheet", (2, 4), 4)  # You can use (2, 4) in place of "D2".
cursor.write_cell(["Method", "metric1", "metric2", "metric3"]) # You can pass list or tuple for multiple writes.
count = 0
for method in ['A', 'B', 'C', 'D', 'E', 'F']:
    cursor.write_cell(method)
    # Run your code
    for i in range(3):
        cursor.write_cell(count)
        count += 1
```
![ex2](https://github.com/CauchyComplete/EasyPyXL/blob/main/images/ex2.png?raw=true)

## Example 3 : move_vertical
```angular2html
import easypyxl
workbook = easypyxl.Workbook("my_excel.xlsx")
cursor = workbook.new_cursor("MySheet", "B1", 5, move_vertical=True)  # move_vertical: Write top to bottom, then move to the next column.
for i in range(25):
    cursor.write_cell(i)
```
![ex3](https://github.com/CauchyComplete/EasyPyXL/blob/main/images/ex3.png?raw=true)

## Example 4 : Multiple cursors, skip_cell()
```angular2html
import easypyxl
workbook = easypyxl.Workbook("my_excel.xlsx")
cursor1 = workbook.new_cursor("Sheet2", "B2", 4)
cursor2 = workbook.new_cursor("Sheet2", "H2", 4, move_vertical=True)
for i in range(100):
    cursor1.write_cell(i)
    cursor2.write_cell(i * 10)
    if i % 5 == 0:
        cursor1.skip_cell(2)  # Skip two cells
```
![ex4](https://github.com/CauchyComplete/EasyPyXL/blob/main/images/ex4.png?raw=true)

## Example 5 : read_cell(), read_line()
```angular2html
import easypyxl
workbook = easypyxl.Workbook("my_excel.xlsx", backup=False)
cursor = workbook.new_cursor("Sheet2", "B3", 4, reader=True)
print(cursor.read_cell(4))
print(cursor.read_line())
print(cursor.read_line(2))
cursor.skip_line(2)
print(cursor.read_cell())
```
outputs:
```angular2html
[2, 3, 4, 5]
[None, None, 6, 7]
[[8, 9, 10, None], [None, 11, 12, 13]]
20
```
