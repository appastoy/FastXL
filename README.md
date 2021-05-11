# FastXL

This project is fast excel file reader C# library. This supports only read. NOT write.

# How to
Excel file(.xlsx, .xlsm) is a set of XML files that are compressed to zip archive.
This library reads XML files in zip archive directly. Especially, For performance, it doesn't use any XML library. 
It uses only regular expression. So you can access minimum things of excel file specification.

# Example

### Load excel file.
```csharp
using FastXL;

// Load book (without all sheets unloaded)
var book = ExcelFile.LoadBook("test.xlsx");

// Load book (with all sheets loaded)
var book = ExcelFile.LoadBook("test.xlsx", loadAllSheet = true);
```

### Access workbook.
```csharp
using FastXL;

var book = ExcelFile.LoadBook("test.xlsx");

// Get sheet by sheet name.
var sheet1 = book["Sheet1"];

// Get sheet by sheet index.
var sheet2 = book[0];

// Iterate sheet list. (read only list)
foreach (var sheet in book.Sheets)
{
  // ...
}
```

### Access worksheet.
```csharp
using FastXL;

var book = ExcelFile.LoadBook("test.xlsx");
var sheet = book["Sheet1"];

// Get sheet name.
var sheetName = sheet.Name;

// Get row count.
var rowCount = sheet.RowCount;

// Get column count.
var columnCount = sheet.ColumnCount;

// Get cell count. (row x column)
var cellCount = sheet.CellCount;

// Is sheet loaded?
var isLoaded = sheet.IsLoaded;

// Load sheet. (It loads once. Second load call do nothing.
sheet.Load();

// Get cell value.
var cellValue = sheet.Cell(row=1, column=2);

```
