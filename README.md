# Examples for FlexCel Studio for .NET

Here you can find all the demos for [FlexCel Studio for .NET](http://www.tmssoftware.com/site/flexcelnet.asp)

You can find a description of each demo in the [documentation](https://doc.tmssoftware.com/flexcel/net/index.html)
**All the demos here are also available when you install FlexCel using the setup.**

**:book: Note** We update this repository automatically every time we release a new FlexCel version. So if you have notifications integrated with github, you can subscribe to this feed to be notified of new releases.


## New in v 7.25 - September 2025


- **New property TExcelFile.OptionsHideObjects lets you set the option to hide objects in a workbook.** The new property [OptionsHideObjects](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.XlsAdapter/XlsFile/OptionsHideObjects.html) allows you to read or write that property of the workbook. The FlexCel renderer will also honor this setting when you export to PDF, HTML, etc.

- **New overload for creating Hyperlinks.** There is a new overload for creating hyperlinks that automatically separates the url from the text mark. (text mark is the part after # in the url)

- **Improved merged cell handling when exporting or printing.** When calculating the Area to print, Excel in some cases might ignore huge merged cells that end up in the last column or row. In this release, we changed our behavior (never ignoring merged cells) to be similar to Excel (ignore merged cells in particular cases)

