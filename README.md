# Examples for FlexCel Studio for .NET

Here you can find all the demos for [FlexCel Studio for .NET](http://www.tmssoftware.com/site/flexcelnet.asp)

You can find a description of each demo in the [documentation](http://www.tmssoftware.biz/flexcel/doc/net/index.html)
**All the demos here are also available when you install FlexCel using the setup.**

**:book: Note** We update this repository automatically every time we release a new FlexCel version. So if you have notifications integrated with github, you can subscribe to this feed to be notified of new releases.


## New on v 6.18.0.0 - December 2017


- **Support for default CryptoAPI xls encrypted files.** Now FlexCel can read and write xls files encrypted with the CryptoAPI encryption. This is the default encryption algorithm for files created by Excel 2003 or newer. With this addition, all modes and encryption algorithms in both xls and xlsx are now supported.

- **Full support for manipulating XML Mappings in xlsx files.** Now XML Mappings will be preserved when opening and saving xlsx/m files, and there are two new commands in the API to set them or read them with code. The new commands are [GetXmlMap](http://www.tmssoftware.biz/flexcel/doc/net/api/FlexCel.Core/ExcelFile/GetXmlMap.html) and [SetXmlMap](http://www.tmssoftware.biz/flexcel/doc/net/api/FlexCel.Core/ExcelFile/SetXmlMap.html) . As usual, APIMate will show how to use `SetXmlMap`. **Note**: The new API only works in xlsx/x files, not xls. Xml mappings inside xls files will still be preserved when opening and saving xls files, but not converted between xls and xlsx.

- **Bug Fix.** Images made transparent with Excel tools might not preserve their transparency when saved as xlsx.

- **Bug Fix.** in .NET Core 2.0 Exceptions thrown by FlexCel would display the message *'Secure binary serialization is not supported on this platform'* instead of the actual error message.

- **Bug Fix.** When rendering shapes with semi-transparent gradients to PDF or SVG the gradients were exported as fully opaque.

- **Bug Fix.** Files with table slicers saved by FlexCel might not open in Excel 2013. (They already worked fine in Excel 2016, and Excel 2010 doesn't support table slicers).

- **Bug Fix.** Rotated shapes inside groups in xlsx files could be rendered wrong.

- **Bug Fix.** Groups that were flipped horizontally or vertically weren't flipped when rendering. Objects inside were flipped, but the groups themselves weren't.

- **Bug Fix.** Filled polygons could be exported wrong to PDF in  some border cases.

- **Bug Fix.** Filled polygons could be exported wrong to images with the SKIA backend used in .NET Core and Android.

- **Bug Fix.** Legacy system colors in drawings inside xls files could be rendered as transparent instead of the correct color in border cases.

- **Bug Fix.** Xlsx files with complex gradients where the stops were not sorted could cause invalid PDF files.

- **Bug Fix.** Textboxes with more than 8224 characters would corrupt the file when saved as xls.

- **Updated SkiaSharp to 1.59.2 for .NET Core.** Now FlexCel will require SkiaSharp 1.59.2 when using it in .NET Core.

