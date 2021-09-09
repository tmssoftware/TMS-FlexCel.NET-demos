# Getting started

A really simple demo on how to create an Excel file with the API.

## Concepts

- Before using FlexCel, you have to add it to the references, and add
  \"**using FlexCel.Core**\" and \"**using FlexCel.XlsAdapter**\" to
  your using statements.


- The most important class here is the [XlsFile](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.XlsAdapter/XlsFile/index.html) class, from where
  you can read and write to any Excel 2 or newer file.

- To set the value for a cell, use [XlsFile.SetCellValue](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.XlsAdapter/XlsFile/SetCellValue.html). You can
  set any kind of object here, not just text. If you set it to
  a [TFormula](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.Core/TFormula/index.html) object, you will enter a formula.

- As explained in the [FlexCel API Developer Guide](https://doc.tmssoftware.com/flexcel/net/guides/api-developer-guide.html), formats in Excel are indexes to an XF (e**X**tended **F**ormat list) 
  To modify the format on a cell, you have to assign an XF index to
  that cell. To create new XF formats, use [XlsFile.AddFormat](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.XlsAdapter/XlsFile/AddFormat.html)
