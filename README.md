# Examples for FlexCel Studio for .NET

Here you can find all the demos for [FlexCel Studio for .NET](http://www.tmssoftware.com/site/flexcelnet.asp)

You can find a description of each demo in the [documentation](http://www.tmssoftware.biz/flexcel/doc/net/index.html)
**All the demos here are also available when you install FlexCel using the setup.**

**:book: Note** We update this repository automatically every time we release a new FlexCel version. So if you have notifications integrated with github, you can subscribe to this feed to be notified of new releases.


## New on v 6.23 - November 2018


- **Updated minimum Required Android version to 8.0 Oreo.** As required by Xamarin and Google Play, now the minimum supported Android version is 8.0 (API Level 26). We removed calls to deprecated methods and now require methods only available in API Level 26 or newer.

- **New methods UnshareWorkbook and IsSharedWorkbook in ExcelFile.** The method [UnshareWorkbook](http://www.tmssoftware.biz/flexcel/doc/net/api/FlexCel.XlsAdapter/XlsFile/UnshareWorkbook.html) allows you to remove all tracking changes from an xls file. (FlexCel doesn't preserve tracking changes in xlsx files). [IsSharedWorkbook](http://www.tmssoftware.biz/flexcel/doc/net/api/FlexCel.XlsAdapter/XlsFile/IsSharedWorkbook.html) allows you to know if an xls file is a shared workbook.

- **New method PivotTableCountInSheet in ExcelFile.** The method [PivotTableCountInSheet](http://www.tmssoftware.biz/flexcel/doc/net/api/FlexCel.XlsAdapter/XlsFile/PivotTableCountInSheet.html) returns the number of pivot tables in the active sheet.

- **Support for calculating function RANK.AVG.** Added support to calculate the Excel function Rank.AVG which was introduced in Excel 2010. See [supported excel functions](http://www.tmssoftware.biz/flexcel/doc/net/about/supported-excel-functions.html#added-functions-in-excel-2010).

- **Now you can find see the call stack in circular formula references when you call RecalcAndVerify.** Now [RecalcAndVerify](http://www.tmssoftware.biz/flexcel/doc/net/api/FlexCel.XlsAdapter/XlsFile/RecalcAndVerify.html) will report the call stack that lead to a cell recursively calling itself, making it simpler for you to track those down in complex spreadsheets. Take a look at the modified [Validate Recalc demo](http://www.tmssoftware.biz/flexcel/doc/net/samples/csharp/api/validate-recalc/index.html) with a file with circular references to see how it works.

- **Bug Fix.** Some xlsx files with legacy headers could fail to load.

- **Bug Fix.** The function IFNA could in very rare corner cases return #N/A if its first parameter was #N/A instead of returning the second parameter.

- **Bug Fix.** There could be an error when copying sheets between workbooks and the sheet copied had a shape with a gradient.

- **Bug Fix.** Floating point numbers that were either infinity or not-a-number were saved wrong in the files and Excel would complain when opening them. Now they will be saved as #NUM! errors. Note that this only happened if you set a cell value explicitly to Double.NAN or Double.Infinity. Formula results which were infinity or nan were already handled fine.

