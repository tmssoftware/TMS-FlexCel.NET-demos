# Examples for FlexCel Studio for .NET

Here you can find all the demos for [FlexCel Studio for .NET](http://www.tmssoftware.com/site/flexcelnet.asp)

You can find a description of each demo in the [documentation](https://doc.tmssoftware.com/flexcel/net/index.html)
**All the demos here are also available when you install FlexCel using the setup.**

**:book: Note** We update this repository automatically every time we release a new FlexCel version. So if you have notifications integrated with github, you can subscribe to this feed to be notified of new releases.


## New in v 7.8 - November 2020


- **.NET 5 Support.** The codebase has been updated to support .NET 5. **Important** :  To use FlexCel in .NET 5, please update to at least this FlexCel version. .NET 5 changed string handing in Windows to use ICU, and that causes problems with older versions of FlexCel in Windows. See [https://github.com/dotnet/runtime/issues/43736](https://github.com/dotnet/runtime/issues/43736)

- **Breaking Change: Removed support for .NET Core 2.0 and 3.0.** We removed support for .NET Core 2.0 and 3.0 as both reached end of life. We keep supporting .NET Core 2.1 and 3.1. See   [https://dotnet.microsoft.com/platform/support/policy/dotnet-core](https://dotnet.microsoft.com/platform/support/policy/dotnet-core)

- **Support for switching Graphics engines in .NET Core or .NET 5.** There is a new property in FlexCelConfig: [GraphicFramework](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.Core/FlexCelConfig/GraphicFramework) which allows you to select between using GDI+ in Windows (native) or SKIA (better compatibility with other platforms which use SKIA)

- **Support for reading fonts from the disk even if the graphics library returns that information.** There is a new property in FlexCelConfig: [ForcePdfFontsFromDisk](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.Core/FlexCelConfig/ForcePdfFontsFromDisk) which allows you to select if FlexCel should use the font data returned by the graphics library if possible, or always scan a folder with fonts.

- **Improved performance in the SKIA graphics backend for .NET 5.** Now SKIA is faster than GDI+ in windows if you use it as the graphics library.

- **Improved compatibility with invalid xls files.** Now FlexCel can ignore invalid ministreams when reading corrupt/invalid xls files.

