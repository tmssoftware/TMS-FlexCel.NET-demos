# Consolidating files

The FlexCel API is oriented to **modifying**
files, instead of reading and creating files as different things. So,
some most important commands on it are [ExcelFile.InsertAndCopyRange](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.Core/ExcelFile/InsertAndCopyRange.html) and
 [ExcelFile.DeleteRange](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.Core/ExcelFile/DeleteRange.html), that copy and delete ranges on existing sheets.

This is a real-world example on how you can use [ExcelFile.InsertAndCopyRange](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.Core/ExcelFile/InsertAndCopyRange.html) to
copy the first sheet of many different Excel files into one big file.

## Concepts

- You can use [ExcelFile.InsertAndCopyRange](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.Core/ExcelFile/InsertAndCopyRange.html) and/or [ExcelFile.InsertAndCopySheets](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.Core/ExcelFile/InsertAndCopySheets.html) to copy ranges
  across different files. Even when it is not as complete as copying
  from the same file, it does copy most of the things.

- [ExcelFile.InsertAndCopyRange](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.Core/ExcelFile/InsertAndCopyRange.html) behaves the same way as Excel. That is, if you
  copy whole rows, the row height and format will be copied, not
  only the values. The same happens with columns, only when copying
  full columns the format and width will be copied to the
  destination. On this demo, we want to copy all Column and Row
  format, so we **have to select the whole sheet**. If we selected a
  smaller range, say (1,1,65535,255) instead of (1,1,65536,256) no
  full column or full row would be selected and not column or row
  format would be copied.

- If the sheets you are copying have formulas or names with references to other files or sheets, you might not get the expected results. You could use [ExcelFile.ConvertFormulasToValues](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.Core/ExcelFile/ConvertFormulasToValues.html) and [ExcelFile.ConvertFormulasToValues](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.Core/ExcelFile/ConvertFormulasToValues.html)
