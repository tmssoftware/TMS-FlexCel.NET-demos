# Printing, previewing and exporting files

FlexCel reporting is oriented to creating files, not to print them. Once you have the files you can
save them, email them or just print them if that's what you really need. But, sometimes you
might want to directly print the report, and here is where
[FlexCelPrintDocument](https://download.tmssoftware.com/flexcel/doc/net/api/FlexCel.Render/FlexCelPrintDocument/index.html) can be helpful. You might want also to export
the report to PDF, and then you would use [FlexCelPdfExport](https://download.tmssoftware.com/flexcel/doc/net/api/FlexCel.Render/FlexCelPdfExport/index.html). Or you might want to export
the Excel file as an image, or to fax it, using [FlexCelImgExport](https://download.tmssoftware.com/flexcel/doc/net/api/FlexCel.Render/FlexCelImgExport/index.html).

All of those components share a common rendering engine, that
\"renders\" the xls file to a canvas, so it can be printed or saved.
Keep in mind that results are not 100% the same, and they cannot be, but
they are very similar.

[FlexCelPrintDocument](https://download.tmssoftware.com/flexcel/doc/net/api/FlexCel.Render/FlexCelPrintDocument/index.html) is a [Winforms PrintDocument](https://docs.microsoft.com/en-us/dotnet/api/system.drawing.printing.printdocument) descendant, and you can use it on the same way. This means you can use it
with a standard [PrintPreviewDialog](https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.printpreviewdialog) or [PrintPreviewControl](https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.printpreviewcontrol) as you
would do with any [PrintDocument](https://docs.microsoft.com/en-us/dotnet/api/system.drawing.printing.printdocument).
