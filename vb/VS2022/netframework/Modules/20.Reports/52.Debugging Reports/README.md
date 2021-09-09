# Debugging reports

As you have seen in the previous demos, tags can get complex once you
start nesting them one inside the other to get a result. This demo shows
how to investigate expressions on your report, and deal with bugs in
your expressions by looking at what is going on inside the hood. Make
sure you read [Debugging Reports](https://doc.tmssoftware.com/flexcel/net/guides/reports-designer-guide.html#debugging-reports) in the Report designer's guide for more in depth information.

## Concepts

- How to use the \<\#debug\> tag in the config sheet to activate debug
  mode. Note that we could get the same effect by setting
  [FlexCelReport.DebugExpressions](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.Report/FlexCelReport/DebugExpressions.html) and [FlexCelReport.ErrorsInResultFile](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.Report/FlexCelReport/ErrorsInResultFile.html).
