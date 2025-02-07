# How to add dropdown list in Google SpreadSheet?
There are many ways to do that, the core concept is to apply a rule with `Citeria` is `DropDown`.

To do that, you can follow one of these ways.

## 1th way (manually)
See [this article](https://github.com/40843245/Google-SpreadSheet/blob/main/How%20to/How%20to%20add%20dropdown%20list%20in%20Google%20SpreadSheet%3F.md)
## 2th way (coding)
Given `range` and `cell` (both are [`Range`](https://developers.google.com/apps-script/reference/spreadsheet/range) instance), 

1. You can create rule through `SpreadsheetApp.newDataValidation().requireValueInRange(range).build();`

2. Then you can apply it through `cell.setDataValidation(rule);` 

### Examples
#### Example 1
For example,

one want to create a dropdown list to range `E15:E20` with options `YES` and `NO`.

<img width="936" alt="image" src="https://github.com/user-attachments/assets/96c2b037-08e4-4517-a913-e132f3beb88a" />

1. One can enter `YES` in `D2`, `NO` in `D3` (for example, you can enter `range` anywhere you want, but the `range` must be consecutive)

2. Type these code, execute `createDropdownListForShowOtherOptions` function in AppScript.

`createDropdownListForShowOtherOptions.gs`

```
function createDropdownListForShowOtherOptions() {
  const range = SpreadsheetApp.getActive().getRange('D2:D3'); // get range that will apply to the rule (`Range` instance).
  const rule = SpreadsheetApp.newDataValidation().requireValueInRange(range).build(); // create a new rule, which applies to Range `range`
  for(let i=15;i<=20;i++){
    const cell = getCell(i,5); // get cell (`Range` instance)
    cell.setDataValidation(rule); // apply rule to the cell.
  }
}
```

`getCell.gs`

```
function getCell(rowIndex,columnIndex) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  return sheet.getRange(rowIndex,columnIndex);
}
```

### output
<img width="124" alt="image" src="https://github.com/user-attachments/assets/3a4d43c8-805e-481f-8974-ba68f703f33d" />

<img width="739" alt="image" src="https://github.com/user-attachments/assets/804f9853-97b4-410c-904a-1c5a7d062488" />

### reference
#### API reference
+ [`Range` class](https://developers.google.com/apps-script/reference/spreadsheet/range)
+ [`DataValidationBuilder` Class](https://developers.google.com/apps-script/reference/spreadsheet/data-validation-builder)
