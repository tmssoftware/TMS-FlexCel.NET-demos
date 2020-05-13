# Creating pdf files with the PDF API

Even when FlexCel is not a full featured PDF package, it does have a
basic PDF API that you can use to create PDF files from scratch.

## Concepts

- How to create a PDF file using FlexCel\'s internal PDF API. The API
  is very similar to GDI+, and allows you to use a PDF \"Canvas\"
  where you can draw things in. To use the API, you need to use the class [PdfWriter](https://download.tmssoftware.com/flexcel/doc/net/api/FlexCel.Pdf/PdfWriter/index.html)

- The PDF API on FlexCel is designed to support exporting Excel
  documents to PDF using **[FlexCelPdfExport](https://download.tmssoftware.com/flexcel/doc/net/api/FlexCel.Render/FlexCelPdfExport/index.html)**. But you can use the
  same API [FlexCelPdfExport](https://download.tmssoftware.com/flexcel/doc/net/api/FlexCel.Render/FlexCelPdfExport/index.html) uses to create your own PDF files with
  code.
