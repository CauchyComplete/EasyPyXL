# EasyPyXL
 This python package is a wrapper of OpenPyXL for easy usage.

## Install
```pip install git+https://github.com/CauchyComplete/EasyPyXL```

## Example 1
```angular2html
import easypyxl
workbook = easypyxl.Workbook("my_excel.xlsx")
cursor = workbook.new_cursor("MySheet", (2, 1), 5)  # New cursor starting from (2, 1), new line every 5 writes.
for i in range(25):
    cursor.write_cell(i)
```
![ex1](https://github.com/CauchyComplete/EasyPyXL/blob/main/images/ex1.png?raw=true)

## Example 2
```angular2html
import easypyxl
workbook = easypyxl.Workbook("my_excel.xlsx")
cursor = workbook.new_cursor("MySheet", (2, 2), 4)
cursor.write_cell(["Method", "TPR", "TNR", "ACC"])  # You can pass list or tuple for multiple writes.
for method in ['A', 'B', 'C', 'D', 'E', 'F']:
    cursor.write_cell(method)
    # Run your code
    for i in range(3):
        cursor.write_cell(i)
```
![ex2](https://github.com/CauchyComplete/EasyPyXL/blob/main/images/ex2.png?raw=true)

## Example 3
```angular2html
import easypyxl
workbook = easypyxl.Workbook("my_excel.xlsx")
cursor = workbook.new_cursor("MySheet", (1, 2), 5, move_vertical=True)  # move_vertical: Write top to bottom, then move to the next column.
for i in range(25):
    cursor.write_cell(i)
```
![ex3](https://github.com/CauchyComplete/EasyPyXL/blob/main/images/ex3.png?raw=true)
