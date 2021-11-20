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

# cellsMerge
## Library overview
ThisWorkbook sheet cells merge, by the grouping target of base column. [Code view here](https://github.com/shoichiros/my-library/blob/master/vba/libraries/CellsMerge.bas)

## Argument description
| Argument        | Description                        |
| --------------- | ---------------------------------- |
| `target_sheet`  | Select Worksheet Object name       |
| `base_column`   | Grouping target of base column     |
| `start_row`     | Data body start row                |
| `target_column` | Merge target column                |
| `is_sum`        | To sum target_column or not to sum |

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

> **WARNING**: Output 2-dimensional arrays is changed Array(Columns, Rows)

## Argument description
| Argument        | Description                      |
| --------------- | -------------------------------- |
| `csv_full_path` | Import target CSV file full path |
| `sql`           | SQL statement                    |

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

> **WARNING**: Be careful with automatic type conversion.
>
> If there is too little data in each column, automatic type conversion will be performed, but the data may not be retrieved correctly. In such a case, please specify the type using CDate() or other methods.

## Argument description
| Argument            | Description                      |
| ------------------- | -------------------------------- |
| `csv_full_path`     | Import target CSV file full path |
| `sql`               | SQL statement                    |
| `paste_start_range` | Data paste start range           |

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