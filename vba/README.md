# How to use VBA libraries

**Contents**
- [How to use VBA libraries](#how-to-use-vba-libraries)
- [cellsMerge](#cellsmerge)
  - [Library overview](#library-overview)
  - [Argument description](#argument-description)
  - [For example code](#for-example-code)
- [CSVImportToArray](#csvimporttoarray)
  - [Library overview](#library-overview-1)
  - [Argument description](#argument-description-1)
  - [For example code](#for-example-code-1)
- [CSVImportToSheet](#csvimporttosheet)
  - [Library overview](#library-overview-2)
  - [Argument description](#argument-description-2)
  - [For example code](#for-example-code-2)
- [sheetImportToArray](#sheetimporttoarray)
  - [Library overview](#library-overview-3)
  - [Argument description](#argument-description-3)
  - [For example code](#for-example-code-3)
- [sheetImportToSheet](#sheetimporttosheet)
  - [Library overview](#library-overview-4)
  - [Argument description](#argument-description-4)
  - [For example code](#for-example-code-4)
- [importAccessToTableSheet](#importaccesstotablesheet)
  - [Library overview](#library-overview-5)
  - [Argument description](#argument-description-5)
  - [For example code](#for-example-code-5)
- [thisworkbookExecuteSQL](#thisworkbookexecutesql)
  - [Library overview](#library-overview-6)
  - [Argument description](#argument-description-6)
  - [For example code](#for-example-code-6)
- [moveFolders](#movefolders)
  - [Library overview](#library-overview-7)
  - [Argument description](#argument-description-7)
  - [For example code](#for-example-code-7)
- [csvFilesMerge](#csvfilesmerge)
  - [Library overview](#library-overview-8)
  - [Argument description](#argument-description-8)
  - [For example code](#for-example-code-8)
- [multipleLayersMkDir](#multiplelayersmkdir)
  - [Library overview](#library-overview-9)
  - [Argument description](#argument-description-9)
  - [For example code](#for-example-code-9)
- [sheetProtection](#sheetprotection)
  - [Library overview](#library-overview-10)
  - [Argument description](#argument-description-10)
  - [For example code](#for-example-code-10)
- [monthLastDay](#monthlastday)
  - [Library overview](#library-overview-11)
  - [Argument description](#argument-description-11)
  - [For example code](#for-example-code-11)
- [makeOutlookMail](#makeoutlookmail)
  - [Library overview](#library-overview-12)
  - [Argument description](#argument-description-12)
  - [For example code](#for-example-code-12)
- [makePDFFile](#makepdffile)
  - [Library overview](#library-overview-13)
  - [Argument description](#argument-description-13)
  - [For example code](#for-example-code-13)
- [makePDFFileAll](#makepdffileall)
  - [Library overview](#library-overview-14)
  - [Argument description](#argument-description-14)
  - [For example code](#for-example-code-14)
- [makePDFWrapSheets](#makepdfwrapsheets)
  - [Library overview](#library-overview-15)
  - [Argument description](#argument-description-15)
  - [For example code](#for-example-code-15)
- [sheetLastRow](#sheetlastrow)
  - [Library overview](#library-overview-16)
  - [Argument description](#argument-description-16)
  - [For example code](#for-example-code-16)
- [wrapPrintOutSheets](#wrapprintoutsheets)
  - [Library overview](#library-overview-17)
  - [Argument description](#argument-description-17)
  - [For example code](#for-example-code-17)

# cellsMerge
## Library overview
ThisWorkbook sheet cells merge, by the grouping target of base column. [Code view here](https://github.com/shoichiros/my-library/blob/master/vba/libraries/CellsMerge.bas)

## Argument description
| Argument      | Description                        |
| ------------- | ---------------------------------- |
| target_sheet  | Select Worksheet Object name       |
| base_column   | Grouping target of base column     |
| start_row     | Data body start row                |
| target_column | Merge target column                |
| is_sum        | To sum target_column or not to sum |

## For example code
```vb
Const BASE_COLUMN As Long = 1
Const START_ROW As Long = 2
Const TARGET_COLUMN As Long = 3
Const IS_SUM As Boolean = True

Dim target_sheet As Worksheet
tatget_sheet = DataSheet

Call cellsMerge(target_sheet, BASE_COLUMN, START_ROW, TARGET_COLUMN, IS_SUM)

```


# CSVImportToArray

## Library overview
Make 2-dimensional arrays after SQL data processing of CSV file. [Code view here](https://github.com/shoichiros/my-library/blob/master/vba/libraries/SQLCSVData.bas)

> <span style="color:red">**WARNING**</span>: Output 2-dimensional arrays is changed Array(Columns, Rows)

> <span style="color:red">**WARNING**</span>: Be careful with automatic type conversion.
>
> If there is too little data in each column, automatic type conversion will be performed, but the data may not be retrieved correctly. In such a case, please specify the type using CDate() or other methods.

## Argument description
| Argument      | Description                      |
| ------------- | -------------------------------- |
| csv_full_path | Import target CSV file full path |
| sql           | SQL statement                    |

## For example code
```vb
Dim csv_full_path As String
Dim sql As String
Dim file_name As String

csv_full_path = Application.GetOpenFilename("CSV(*.csv), *.csv", , "csv")
file_name = Dir(csv_full_path)
sql = "SELECT *" _
    & " FROM [" & file_name & "]"

Dim data_lists As Variant
data_lists = CSVImportToArray(csv_full_path, sql)

```


# CSVImportToSheet

## Library overview
After data processing CSV file with SQL, insert data to Worksheet. [Code view here](https://github.com/shoichiros/my-library/blob/master/vba/libraries/SQLCSVData.bas)

> <span style="color:red">**WARNING**</span>: Be careful with automatic type conversion.
>
> If there is too little data in each column, automatic type conversion will be performed, but the data may not be retrieved correctly. In such a case, please specify the type using CDate() or other methods.

## Argument description
| Argument          | Description                      |
| ----------------- | -------------------------------- |
| csv_full_path     | Import target CSV file full path |
| sql               | SQL statement                    |
| paste_start_range | Data paste start range           |

## For example code
```vb
Dim csv_full_path As String
Dim sql As String
Dim file_name As String
Dim paste_start_range As Range

csv_full_path = Application.GetOpenFilename("CSV(*.csv), *.csv", , "csv")
file_name = Dir(csv_full_path)
sql = "SELECT *" _
    & " FROM [" & file_name & "]"

Set paste_start_range = DataSheet.Range("A2")

Call CSVImportToSheet(csv_full_path, sql, paste_start_range)

```


# sheetImportToArray

## Library overview
After data processing ThisWorkBook sheet with SQL, output 2-dimensional arrays data. [Code view here](https://github.com/shoichiros/my-library/blob/master/vba/libraries/SQLSheetData.bas)

> <span style="color:red">**WARNING**</span>: Output 2-dimensional arrays is changed Array(Columns, Rows)

> <span style="color:red">**WARNING**</span>: Be careful with automatic type conversion.
>
> If there is too little data in each column, automatic type conversion will be performed, but the data may not be retrieved correctly. In such a case, please specify the type using CDate() or other methods.

## Argument description
| Argument | Description   |
| -------- | ------------- |
| sql      | SQL statement |

## For example code
```vb
Dim sheet_name As String
Dim sql As String

sheet_name = Sheet1.Name
sql = "SELECT *" _
    & " FROM [" & sheet_name & "$]"

Dim data_lists As Variant
data_lists = sheetImportToArray(sql)

```


# sheetImportToSheet

## Library overview
After data processing ThisWorkBook sheet with SQL, paste another sheet. [Code view here](https://github.com/shoichiros/my-library/blob/master/vba/libraries/SQLSheetData.bas)

> <span style="color:red">**WARNING**</span>: Be careful with automatic type conversion.
>
> If there is too little data in each column, automatic type conversion will be performed, but the data may not be retrieved correctly. In such a case, please specify the type using CDate() or other methods.

## Argument description
| Argument          | Description            |
| ----------------- | ---------------------- |
| sql               | SQL statement          |
| paste_start_range | Data paste start range |

## For example code
```vb
Dim sheet_name As String
Dim sql As String
Dim paste_start_range As Range

sheet_name = DataSheet.Name
sql = "SELECT *" _
    & " FROM [" & sheet_name & "$]"
Set paste_start_range = DataSheet.Range("A2")

Call sheetImportToSheet(sql, paste_start_range)

```


# importAccessToTableSheet

## Library overview
After data processing Access table or query with SQL, paste ThisWorkbook table data. [Code view here](https://github.com/shoichiros/my-library/blob/master/vba/libraries/ExtractionAccessData.bas)

## Argument description
| Argument    | Description                  |
| ----------- | ---------------------------- |
| sql         | SQL statement                |
| db_path     | Target Access file full path |
| paste_sheet | Objectsheet name             |
| is_table    | Make Data table or not table |

## For example code
```vb
Dim sql As string
Dim db_path As string
Dim paste_sheet As string
Dim is_table As Boolean

' Olso sql = "SELECT name, age FROM [queryName]"
sql = "SELECT name, age FROM [dataTable]"
db_path = "C:\Users\{Your Username}\Desktop\AccessData.accdb"
paste_sheet = DataSheet
is_table = True

Call importAccessToTableSheet(sql, db_path, paste_sheet, is_table)

```


# thisworkbookExecuteSQL

## Library overview
UPDATE and INSERT INTO SQL statement execute in ThisWorkbook. [Code view here](https://github.com/shoichiros/my-library/blob/master/vba/libraries/ExecuteSQL.bas)

> <span style="color:red">**WARNING**</span>: Not use Delete SQL statement.

## Argument description
| Argument | Description   |
| -------- | ------------- |
| sql      | SQL statement |

## For example code
```vb
Dim sql As String

' Olso sql = "INSERT INTO [dataTableName] VALUES ('joy', 26)"
sql = "UPDATE [DataSheet.Name$] SET name = 'Joy' WHERE name = 'j'"

Call thisworkbookExecuteSQL(sql)

```


# moveFolders

## Library overview
The local folders or network folders into move another folder. [Code view here](https://github.com/shoichiros/my-library/blob/master/vba/libraries/DirOperation.bas)

## Argument description
| Argument           | Description              |
| ------------------ | ------------------------ |
| before_folder_path | Move target folder path  |
| after_folder_path  | Moved target folder path |

## For example code
```vb
Dim before_folder_path As String
Dim after_folder_path As String

before_folder_path = "C:\Users\{Your username}\Desktop\test_folder"
after_folder_path = "C:\Users\{Your username}\Desktop\test\"

Call moveFolders(before_folder_path, after_folder_path)

```


# csvFilesMerge

## Library overview
If same Columns in CSV, all merge CSV files in target folder. [Code view here](https://github.com/shoichiros/my-library/blob/master/vba/libraries/MultipleCSVMerge.bas)

## Argument description
| Argument      | Description                                  |
| ------------- | -------------------------------------------- |
| target_folder | Folder path with CSV files you want to merge |
| output_folder | Folder path merged CSV files are stored      |

## For example code
```vb
Const TARGET_FOLDER As String = "C:\Users\{Your username}\Desktop\test\"
Const OUTPUT_FOLDER As String = "C:\Users\{Your username}\Desktop\"

Call csvFilesMerge(TARGET_FOLDER, OUTPUT_FOLDER)

```


# multipleLayersMkDir

## Library overview
The make new multiple layers folder, because default VBA MkDir function is not making multiple layers folder. [Code view here](https://github.com/shoichiros/my-library/blob/master/vba/libraries/MultipleMkdir.bas)

## Argument description
| Argument           | Description                     |
| ------------------ | ------------------------------- |
| output_folder_path | You want to make directory path |

## For example code
```vb
Const OUTPUT_FOLDER_PATH As String = "C:\Users\{Your username}\Desktop\tests\test"

Call multipleLayersMkDir(OUTPUT_FOLDER_PATH)

```


# sheetProtection

## Library overview
ThisWorkbook sheet or multiple sheets protection. [Code view here](https://github.com/shoichiros/my-library/blob/master/vba/libraries/Protection.bas)

## Argument description
| Argument      | Description                        |
| ------------- | ---------------------------------- |
| is_protect    | Sheet protect or not Sheet protect |
| password      | Sheet password                     |
| target_sheets | Protect target Sheets array        |

## For example code
```vb
Const PASSWORD As String = "TestPassword"

Dim sheet_lists As Variant
sheet_lists = Array(DataSheet, OutputSheet)

Call sheetProtection(True, PASSWORD, sheet_lists) ' Protect
Call sheetProtection(False, PASSWORD, sheet_lists) ' Unprotect
```


# monthLastDay

## Library overview
Get the last day of the specified month. [Code view here](https://github.com/shoichiros/my-library/blob/master/vba/libraries/LastDayFunc.bas)

## Argument description
| Argument     | Description                                                                                      |
| ------------ | ------------------------------------------------------------------------------------------------ |
| month_number | 0 is this month. 1 is next month. -1 is the previous month. ... -3, -2, -1, 0, 1, 2, 3 ...       |
| is_character | Output is String hyphen "2021-03-31" the ID, File name olso Folder name or Date slash 2021/03/31 |

## For example code
```vb
Const MONTH_NUMBER As Long = 0
Const IS_CHARACTER As Boolean = True

Dim date_data As String
date_data = monthLastDay(MONTH_NUMBER, IS_CHARACTER) ' "2021-03-31"

```


# makeOutlookMail

## Library overview
The make new Outlook email, to save the draft folder. [Code view here](https://github.com/shoichiros/my-library/blob/master/vba/libraries/MakeOutlook.bas)

> <span style="color:blue">**Required**</span>: Microsoft Outlook 16.0 Object Library

## Argument description
| Argument                          | Description                                         |
| --------------------------------- | --------------------------------------------------- |
| address                           | To address                                          |
| subject                           | Subject                                             |
| body                              | Mail Body                                           |
| is_attach                         | Whether attachments files exist or not.             |
| attach_file_path_array (optional) | is_attach = True is required attach_file_path_array |

## For example code
```vb
Const ADDRESS As String = "example@example.com"
Const SUBJECT As String = "example subject"
Const BODY As String = "example contents"
Const IS_ATTACH As Boolean = True

Const first_attach_file As String = "C:\Users\{Your username}\Desktop\test\test-contents.txt"
Const second_attach_file As String = "C:\Users\{Your username}\Desktop\test\test.txt"

Dim attach_file_path_array As Variant
attach_file_path_array = Array(first_attach_file, second_attach_file)

Call makeOutlookMail(ADDRESS, SUBJECT, BODY, IS_ATTACH, attach_file_path_array)

```


# makePDFFile

## Library overview
Some sheet choice, make PDF files in ThisWorkbook folder. [Code view here](https://github.com/shoichiros/my-library/blob/master/vba/libraries/MakePDF.bas)

## Argument description
| Argument           | Description                                     |
| ------------------ | ----------------------------------------------- |
| sheets_name_array  | The `SheetObject.Name` or sheets name, to Array |
| output_folder_name | Export folder name                              |

## For example code
```vb
Dim sheets_name_array As Variant
sheets_name_array = Array(DataSheet.Name, OutPutSheet.Name)

Const OUTPUT_FOLDER_NAME As String = "export_folder"

Call makePDFFile(sheets_name_array, OUTPUT_FOLDER_NAME)

```


# makePDFFileAll

## Library overview
ThisWorkbook all sheets, make only one PDF file in ThisWorkbook folder. [Code view here](https://github.com/shoichiros/my-library/blob/master/vba/libraries/MakePDF.bas)

## Argument description
| Argument           | Description        |
| ------------------ | ------------------ |
| output_file_name   | Export file name   |
| output_folder_name | Export folder name |

## For example code
```vb
Const OUTPUT_FILE_NAME As String = "export_file"
Const OUTPUT_FOLDER_NAME As String = "export_folder"

Call makePDFFileAll(OUTPUT_FILE_NAME, OUTPUT_FOLDER_NAME)

```


# makePDFWrapSheets

## Library overview
From some sheets, make only one PDF file in ThisWorkbook folder. [Code view here](https://github.com/shoichiros/my-library/blob/master/vba/libraries/MakePDF.bas)

## Argument description
| Argument           | Description                                     |
| ------------------ | ----------------------------------------------- |
| sheets_name_array  | The `SheetObject.Name` or sheets name, to Array |
| output_file_name   | Export file name                                |
| output_folder_name | Export folder name                              |

## For example code
```vb
Dim sheets_name_array As Variant
sheets_name_array = Array(DataSheet.Name, OutPutSheet.Name)

Const OUTPUT_FILE_NAME As String = "export_file"
Const OUTPUT_FOLDER_NAME As String = "export_folder"

Call makePDFWrapSheets(sheets_name_array, OUTPUT_FILE_NAME, OUTPUT_FOLDER_NAME)

```


# sheetLastRow

## Library overview
Get taget sheet last row. [Code view here](https://github.com/shoichiros/my-library/blob/master/vba/libraries/SheetLastRow.bas)

## Argument description
| Argument                | Description                                     |
| ----------------------- | ----------------------------------------------- |
| target_sheet            | The last row search target sheet                |
| check_column (optional) | The last row search column, default column is 1 |

## For example code
```vb
Dim target_sheet As Worksheet
target_sheet = DataSheet

Const CHECK_COLUMN As Long = 2

Call sheetLastRow(target_sheet, CHECK_COLUMN)

```


# wrapPrintOutSheets

## Library overview
ThisWorkbook sheets to all print out. [Code view here](https://github.com/shoichiros/my-library/blob/master/vba/libraries/WrapPrintOutSheet.bas)

## Argument description
| Argument | Description |
| -------- | ----------- |
| none     | none        |

## For example code
```vb
Call wrapPrintOutSheets()

```