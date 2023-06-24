# Examples for FlexCel Studio for .NET

Here you can find all the demos for [FlexCel Studio for .NET](http://www.tmssoftware.com/site/flexcelnet.asp)

You can find a description of each demo in the [documentation](https://doc.tmssoftware.com/flexcel/net/index.html)
**All the demos here are also available when you install FlexCel using the setup.**

**:book: Note** We update this repository automatically every time we release a new FlexCel version. So if you have notifications integrated with github, you can subscribe to this feed to be notified of new releases.


## New in v 7.17 - June 2023


- **Optimized support for .NET 7.** We include now packages compiled against .NET 7

- **SkiaSharp updated to 2.88.3.** The minimum SkiaSharp now required is 2.88.3.

- **Removed support for .NET Core 3.1.** As .NET Core 3.1 reached EOL, now the minimum .NET Core version supported is 5.0

- **Improved API for defining columns in tables.** Now you can define a totals formula or a column formula for the columns in the table, if needed. As usual, APIMate will show you how to do it.

- **Now FlexCel preserves digital signatures in macros.** When you have digitally signed macros in a file, now FlexCel will preserve them when opening and saving that file

- **New property  DeleteEmptyBandsFixed in FlexCelReport controls what to do with empty fixed bands.** The new function [DeleteEmptyBandsFixed ](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.Report/FlexCelReport/DeleteEmptyBandsFixed.html) lets you define what happens if a fixed band has zero records.

- **Bug Fix.** When recreating a table by calling AddTable and SetTable, the cell references could become invalid

- **Bug Fix.** The functions IFERROR, ISERROR and ISERR could sometimes return the error instead of the result of the function.

- **Bug Fix.** The functions COUNTIF, SUMIF and similar could behave wrong in some cases where you used wildcards. See [https://support.tmssoftware.com/t/countif-formula-with-wildcard-failing-after-recalc-method-is-called/19266](https://support.tmssoftware.com/t/countif-formula-with-wildcard-failing-after-recalc-method-is-called/19266)

- **Bug Fix.** The function [TRIM](https://support.microsoft.com/en-us/office/trim-function-410388fa-c5df-49c6-b16c-9e5630b479f9) in Excel removes double spaces in the middle of a text, while FlexCel's implementation would remove only spaces at the beginning at end. Also Excel's TRIM only removes spaces (character 32) and not other whitespace like tabs. FlexCel's implementation now does the same.

- **Bug Fix.** The function =NUMBERVALUE() could throw an Exception in some border cases

