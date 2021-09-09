# Reading Excel files

A demo showing how to read the contents of an Excel file using FlexCel.

## Concepts

- To read an Excel file you use the [XlsFile](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.XlsAdapter/XlsFile/index.html) class, from where you
  can read and write to any Excel 2.0 or newer
  file.

- To get the value for a single cell, use [XlsFile.GetCellValue](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.XlsAdapter/XlsFile/GetCellValue.html).

- To get the value for a cell when looping a full sheet, use
  [XlsFile.GetCellValueIndexed](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.XlsAdapter/XlsFile/GetCellValueIndexed.html). It is faster than using
  GetCellValue since you will only read the used cells.

- [XlsFile.GetCellValue](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.XlsAdapter/XlsFile/GetCellValue.html) and [XlsFile.GetCellValueIndexed](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.XlsAdapter/XlsFile/GetCellValueIndexed.html) can return one of the following
  objects:

  - null

  - Double

  - Boolean

  - String

  - [TRichString](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.Core/TRichString/index.html)

  - [TFormula](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.Core/TFormula/index.html)

  - 



- With GetCellValue and GetCellValueIndexed you will get the actual
  values. But if you want to actually display formatted data (for
  example if you have the number 2 with 2 decimals, and you want to
  display 2.00 instead of 2), you need to use other methods. There
  are 2 ways to do it:

   1. [XlsFile.GetStringFromCell](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.XlsAdapter/XlsFile/GetStringFromCell.html) will return a rich string with the
   cell formatted.

   2. FormatValue will format an object with a
   specified format and then return the corresponding rich string.
   TFlxNumberFormat.FormatValue is used internally by
   GetStringFromCell.

- In Excel, **Dates are doubles**. The only difference between a date
  and a double is on the format on the cell. With
  FormatValue you can get the actual string that is
  displayed on Excel. Also, to convert this double to a DateTime,
  you can use [FlxDateTime.FromOADate](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.Core/FlxDateTime/FromOADate.html).
