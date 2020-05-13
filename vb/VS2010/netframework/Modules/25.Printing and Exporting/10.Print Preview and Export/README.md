# Printing, previewing and exporting

A small demo on how to use [FlexCelPrintDocument](https://download.tmssoftware.com/flexcel/doc/net/api/FlexCel.Render/FlexCelPrintDocument/index.html) to natively print
and preview any existing Excel File, and on how to use
[FlexCelImgExport](https://download.tmssoftware.com/flexcel/doc/net/api/FlexCel.Render/FlexCelImgExport/index.html) to convert any Excel file to images. Pdf
export is not shown here, but it is on other examples.

## Concepts

- [FlexCelPrintDocument](https://download.tmssoftware.com/flexcel/doc/net/api/FlexCel.Render/FlexCelPrintDocument/index.html) output is not 100% identical to Excel output,
  and it can\'t be that way. But it is very similar, and this
  includes fonts, colors, margins, headers/footers/images, etc. It
  can print cells with multiple fonts, it can replace the macros on
  headers and footers (like \"&CPage &P of &N\"), it can show
  conditional formats, and the list goes on.

- You can customize the final output by assigning a PagePrint event.
  On this demo, it is customized so it prints \"Confidential\" on
  each page. (if you select the checkbox)

- How to export the generated output to images or a multipage tiff, 
  which is the format used to send faxes. On
  this example, the tiff quality is hard-coded on the demo to 96
  DPI, but you can increase to get better-looking pictures (and
  higher file sizes)
