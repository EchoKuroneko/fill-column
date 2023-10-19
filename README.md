# FillColumn AppScript Function

This script provides a custom functionality to fill values from a given range in Google Sheets.

It fills values repetitively vertically and in sequence to return a larger array.

## Purpose

This custom function was created to simplify the process of filling repetitive values in a single column. Writing repetitive values manually or using lengthy formula and applying them to each row can be time-consuming and error-pron, such as using `If()` function is viable, where we can provide some sort of logic and if the logic is fullfilled then to write certain data otherwise write elseif condition for it.

The `FillColumn` function enables users to automate this process, enhancing efficiency and streamlining data filling.

### Built-in Formula

  ```
  =If(logical_expression, value_if_true, value_if_false)
  ```

- Example:
    
    Suppose there is a `helper column` containing data to be filled with a specified number of `repetitions`. The following formula can be used:

    ```
    =IF(row_count <= repition, fill_this_data_when_true, or_use_another_recursive_if_condition_when_false)
    ```
    
    For simplicity, suppose the repetition of value is 10 And there already exist helper Column `D` upto 5 rows

    ```
    =IF(ROW(B1)<=10, D$1, IF(ROW(B1)<=10*2, D$2, IF(ROW(B1)<=10*3, D$3, IF(ROW(B1)<=10*4, D$4, IF(ROW(B1)<=10*5, D$5, "")))))
    ```


## Usage

### Function

```javascript
FillColumn(dataRange,repetition,sheetName)
```

### Parameters

- dataRange (${\color{lime}required}$) : Range of the data set that will be used to fill i.e `helper column` (e.g., "B1:B10")

  ${\color{lightblue}Input}$ : `string`

- repetition (${\color{lime}required}$) : Number of times to repeat the same value
    
    ${\color{lightblue}Input}$ : `integer`

- sheetName (${\color{orange}optional}$) : Name of the sheet to be used.

    ${\color{lightgreen}Default}$ : Active Sheet
    
> [!NOTE]
> The range parameter must be provided as a string within the formula and the function uses the active sheet if no sheet name is provided.

### Example

```
=FillColumn("B1:B10", 10, "Sheet1")
```
Or,
```
=FillColumn("B1:B10", 10)
```