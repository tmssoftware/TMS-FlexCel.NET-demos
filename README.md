# Examples for FlexCel Studio for .NET

Here you can find all the demos for [FlexCel Studio for .NET](http://www.tmssoftware.com/site/flexcelnet.asp)

You can find a description of each demo in the [documentation](https://doc.tmssoftware.com/flexcel/net/index.html)
**All the demos here are also available when you install FlexCel using the setup.**

**:book: Note** We update this repository automatically every time we release a new FlexCel version. So if you have notifications integrated with github, you can subscribe to this feed to be notified of new releases.


## New in v 7.11 - September 2021


- **Support for .NET 6 and Visual Studio 2022.** Official support for .NET 6 and Visual Studio 2022

- **New overloads of methods for getting image information that take objectIndexes instead of imageIndexes.** The methods [GetImageProperties](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.Core/ExcelFile/GetImageProperties.html), [GetImageName](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.Core/ExcelFile/GetImageName.html), [SetImageProperties](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.Core/ExcelFile/SetImageProperties.html), [DeleteImage](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.Core/ExcelFile/DeleteImage.html) and [ClearImage](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.Core/ExcelFile/ClearImage.html) now have overloads taking an objectIndex/objectPath pair instead of an imageIndex.

- **New methods to convert between imageindexes and objectindexes with support for grouped shapes.** The new methods [ImageIndexToObjectPath](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.Core/ExcelFile/ImageIndexToObjectPath.html) and [ObjectPathToImageIndex](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.Core/ExcelFile/ObjectPathToImageIndex.html) can convert between imageIndexes and objectIndexes taking care of nested objects. The existing methods would only work in non-grouped objects.

- **You can now read and write the links of a camera object.** The methods [GetShapeLinkedCell](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.Core/ExcelFile/GetShapeLinkedCell.html) and [SetShapeLinkedCell](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.Core/ExcelFile/SetShapeLinkedCell.html) now work in picures too, allowing to read or create camera objects besides to shapes with their text linked.

- **Improved loading of Excel 3, 4 and 95 files.** Now FlexCel can load camera objects in Excel 3, 4 and 95 files. It will also load the image names for images in Excel 95 files (Excel 4 and older don't store an image name)

- **SkiaSharp updated to 2.80.3.** The SkiaSharp library was updated from 2.80.2 to 2.80.3.

