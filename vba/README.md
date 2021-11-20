# How to use VBA libraries

- [How to use VBA libraries](#how-to-use-vba-libraries)
- [CellsMerge](#cellsmerge)
  - [Library overview](#library-overview)
  - [Argument description](#argument-description)
  - [For example code](#for-example-code)

# CellsMerge

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
```Visual Basic
Const BASE_COLUMN As Long = 1
Const START_ROW As Long = 2
Const TARGET_COLUMN As Long = 3
Const IS_SUM As Boolean = True

Dim target_sheet As Worksheet
tatget_sheet = DataSheet

Call cellsMerge(target_sheet, BASE_COLUMN, START_ROW, TARGET_COLUMN, IS_SUM)

```