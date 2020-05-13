# Flexcel object explorer

Shapes in an Excel file can be nested in complex ways. For example, you
might have a \"Group\" shape that inside has a \"Picture\" shape and
other \"Group\" shape with other pictures inside. You can access the
whole hierarchy of objects with [ExcelFile.GetObjectProperties](https://download.tmssoftware.com/flexcel/doc/net/api/FlexCel.Core/ExcelFile/GetObjectProperties.html) but it
can be hard to visualize. This application shows a tree view of all the
objects in a sheet, and can be useful when trying to understand the
structure of the objects in an xls file.

## Concepts

- How to use the API to read the objects and properties in a file.

- How to use [ExcelFile.RenderObject](https://download.tmssoftware.com/flexcel/doc/net/api/FlexCel.Core/ExcelFile/RenderObject.html) to visually render the object
  (as shown in the preview pane). Internally, RenderObject uses
  [ExcelFile.GetObjectProperties](https://download.tmssoftware.com/flexcel/doc/net/api/FlexCel.Core/ExcelFile/GetObjectProperties.html) in the same way as this application,
  and draws the properties into a .NET image.

- Custom properties (The ones shown in the grid at the top) can hold
  integers, booleans, strings or other datatypes. You need to know
  of which type is the property you are trying to read. Since
  FlexCel does not know it, it will try to show it as a string or as
  an integer.
