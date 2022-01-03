## ExcelPeek v0.2 - User Guide

----

## Note on Using This Tool

The library that ExcelPeek relies on to interface with a spreadsheet, openpyxl, requires that files be in the ".xlsx" format.  For 99% of use-cases, all this means is that if the file is another type of Excel file (".xls", ".xlsb", ".xlsm", etc.), then all one needs to do is open the file in something like LibreOffice Calc (DO NOT open the file in Excel on a Windows-based system!!!) and save the file as type ".xlsx".  This will still keep any macros intact and allow investigation to be undertaken.

## Simple Examples For Getting Started

The ExcelPeek help screen can be shown by typing the following:

```bash
<install directory>/excelpeek/excelpeek.py --help
```

All flags must be preceded by 2 hyphens "--".

## Command-line Usage Information

ExcelPeek has several command-line flags which are described below.

| Flag | Description |
|------------|-------------|
| --file | Required - file being examined. |
| --analyze | Optional - list the particulars of all worksheets. |
| --list | Optional - list sheets |
| --sheet | Optional - pick the sheet you want to investigate. |
| --sheetstats | Optional - list the particulars of a specific sheet. |
| --range | Optional - range of cells to investigate - e.g. A1:Z200. |
| --dump  | Optional - dump any non-empty cells. |
| --debug | Optional - Prints verbose logging to the screen to troubleshoot issues. |
| --help | Optional - Prints list of flags. |

## Use Cases
Analyzing all worksheets in an Excel file:
```
<install directory>/excelpeek/excelpeek.py --file d3e6387c8c6ba3060a1dd90333a80ccac9c5dee328770f22e6860d314d8eb547.xlsx --analyze
[*] Length Arguments: 4
[*] Arguments: 
file: d3e6387c8c6ba3060a1dd90333a80ccac9c5dee328770f22e6860d314d8eb547.xlsx
analyze: True

****************************************************************************************************
0. Info
[*] Max Row: 4
[*] Found 4 rows of data.
[*] Max Column: 10
[*] Found 10 columns of data.
[*] Sheet State: visible
****************************************************************************************************
1. Macro1
[*] Max Row: 275
[*] Found 275 rows of data.
[*] Max Column: 19
[*] Found 19 columns of data.
[*] Sheet State: hidden
****************************************************************************************************
2. Sheet1
[*] Max Row: 4976
[*] Found 4976 rows of data.
[*] Max Column: 82
[*] Found 82 columns of data.
[*] Sheet State: hidden
****************************************************************************************************

[*] Program Complete

```


Listing all the sheets in an Excel file:
```bash
<install directory>/excelpeek/excelpeek.py --file index-280403604.xlsx --list
[*] Length Arguments: 4
[*] Arguments: 
file: index-280403604.xlsx
list: True

[*] Worksheets available are:

0. Sheet
1. Sbur1
2. Sbur2
3. Sbur3
4. Kon
5. DEFW3
6. DEFW2
7. DEFW
8. Beff1
9. Beff2
10. Beff3
11. Beff4
12. Beff5
13. Beff6
14. Beff7
15. Beff8
```

Based on the above example, get the stats on a single sheet:
```bash
<install directory>/excelpeek/excelpeek.py --file index-280403604.xlsx --sheet Sbur1 --sheetstats
[*] Length Arguments: 6
[*] Arguments: 
file: index-280403604.xlsx
sheet: Sbur1
sheetstats: True

[*] Loading worksheet: Sbur1
[*] Max Row: 84
[*] Found 84 rows of data.

[*] Program Complete
```

Based on the above examples, dump the data held within a single sheet:
```bash
<install directory>/excelpeek/excelpeek.py --file index-280403604.xlsx --sheet Sbur1 --dump
[*] Length Arguments: 6
[*] Arguments: 
file: index-280403604.xlsx
sheet: Sbur1
dump: True

[*] Loading worksheet: Sbur1
[-] Output name not specified.  Using default: Sbur1.txt
[*] Output of data rows...
Value: =CHAR(210-100)
Value: =CHAR(86-1)
Value: =CHAR(120-53)
Value: =CHAR(200-135)
Value: =CHAR(182-100)
Value: =CHAR(102-54)
Value: =CHAR(179-110)
Value: =CHAR(215-99)
Value: =CHAR(100-25)
Value: =CHAR(210-101)
Value: =CHAR(193-110)
Value: =CHAR(222-101)
Value: =CHAR(103-50)
Value: =CHAR(208-100)
Value: =CHAR(200-89)
Value: =CHAR(203-100)
Value: =CHAR(201-100)
Value: =CHAR(200-103)
Value: =CHAR(104-46)
Value: =CHAR(216-99)
Value: =CHAR(205-101)
Value: =CHAR(186-110)
Value: =CHAR(215-101)
Value: =CHAR(100-32)
Value: =CHAR(219-101)
Value: =CHAR(205-100)
Value: =CHAR(170-96)
Value: =CHAR(220-100)
Value: =CHAR(102-51)
Value: =CHAR(201-102)
Value: =CHAR(213-128)
Value: =CHAR(101-51)
Value: =CHAR(219-100)
Value: =CHAR(176-110)
Value: =CHAR(213-101)
Value: =CHAR(200-85)
Value: =CHAR(201-101)
Value: =CHAR(164-80)
Value: g
Value: =CHAR(160-90)
Value: o
Value: o
Value: d
Value: "
Value: ,
Value: \
Value: .
Value: 1
Value: "
Value: :
Value: .
Value: ,
Value: \
Value: =_xlfn.ARABIC("CXI")
Value: =_xlfn.ARABIC("LXXVI")
Value: =_xlfn.ARABIC("LXV")
Value: =_xlfn.ARABIC("LXVII")
Value: =_xlfn.ARABIC("CI")
Value: =_xlfn.ARABIC("LXI")
Value: R
Value: =_xlfn.ARABIC("CXIV")
Value: E
Value: T
Value: U
Value: (
Value: )
Value: N
Value: =
Value: C
Value: A

[*] Program Complete
```
