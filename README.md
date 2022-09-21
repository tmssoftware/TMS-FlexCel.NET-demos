# Examples for FlexCel Studio for .NET

Here you can find all the demos for [FlexCel Studio for .NET](http://www.tmssoftware.com/site/flexcelnet.asp)

You can find a description of each demo in the [documentation](https://doc.tmssoftware.com/flexcel/net/index.html)
**All the demos here are also available when you install FlexCel using the setup.**

**:book: Note** We update this repository automatically every time we release a new FlexCel version. So if you have notifications integrated with github, you can subscribe to this feed to be notified of new releases.


## New in v 7.16 - September 2022


- **Support for using .NET 6 directly in iOS and Android, instead of Xamarin.** Now we support the new multi-target packages in .NET 6, which replace Xamarin.

- **Breaking Change: The default for the GraphicFramework property is now Native.** The property  [GraphicFramework](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.Core/FlexCelConfig/GraphicFramework.html) now defaults to use the native framework if not manually specified. The native graphic frameworks normally work better than Skia, and don't require SkiaSharp.

- **Support for using native graphics engine in .NET 6 for Android, macOS and iOS.** In Nov 2020 [FlexCel 7.8](https://doc.tmssoftware.com/flexcel/net/about/whatsnew.html#new-in-v-78---november-2020) added support for switching graphics engines in Windows. Now when in .NET 6 you can also switch graphics engines in macOS, iOS (you can use CoreGraphics or Skia) and Android (you can use Native or Skia). To select between native or SkiaSharp in any platform, use the property  [GraphicFramework](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.Core/FlexCelConfig/GraphicFramework.html)

- **SkiaSharp updated to 2.88.1.** The minimum SkiaSharp now required is 2.88.1. This was needed to support multi-targeting in .NET 6

- **Support for different numeric systems in cell formatting.** Now if you format a cell with a different numeric system like for example "[$-2000000]#,##0.00", FlexCel will render those numbers correctly. (see [https://ansarichat.wordpress.com/2018/02/20/how-to-type-arabic-numerals-in-excel/](https://ansarichat.wordpress.com/2018/02/20/how-to-type-arabic-numerals-in-excel/) )

- **Bug Fix.** When rendering charts that used =Offset to define the data, and some columns or rows were hidden, FlexCel could fail to hide the values when the chart was set to not plot hidden cells.

- **Bug Fix.** In some rare cases when there was merged cells whose first cell was hidden the background might not be exported to pdf.

- **Bug Fix.** If printing gridlines and there were hidden columns or rows, the gridlines could be printed over the real borders of a cell.

- **Bug Fix.** When exporting to CSV, there could be errors if you manually set cell values to NaN.

- **Bug Fix.** If exporting to PDF and the "normal" font of the spreadsheet was Calibri 9 columns could be wider than expected.

- **Bug Fix.** FlexCel could hang while loading some invalid third-party files.

