# Templates in the executable

Sometimes you might want to distribute your templates embedded in your application,
instead of shipping the separately.

## Concepts

- How to embed a template directly in the exe file. For
  step-by-step instructions please read [Embedding Excel files in your application](https://doc.tmssoftware.com/flexcel/net/tips/embedding-excel-files-in-your-application.html) 


- Dealing with **includes**. Included files are normally searched on
  the same path as the original file, but here, as we are reading
  from a stream, we must tell FlexCel where to find the embedded
  template. This is done with the [FlexCelReport.GetInclude](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.Report/FlexCelReport/GetInclude.html) event.

- While here we show how to use a stream to read a template from the
  executable, you can use a very similar approach to store your
  templates in a database, or in any place you like. Just call the
  [FlexCelReport.Run](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.Report/FlexCelReport/Run.html) with a stream with your data, and
  assign the [FlexCelReport.GetInclude](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.Report/FlexCelReport/GetInclude.html) event to assign the data for the included
  files.
