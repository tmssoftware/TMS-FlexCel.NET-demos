# Using Linq as datasource

**THIS DEMO NEEDS .NET 3.5 OR NEWER.**

Most of the demos here use datasets as datasources. This is just for
convenience, because Linq is not supported in .NET 2.0 and so if we used
Linq those demos wouldn\'t work for everybody, and also because the
focus is in the Excel templates, not so much in the data layer. But you
can use any IQueryable\<T\> collection as a datasource in a FlexCel
report, and this is what we will show here.

## Concepts

- How to run a report from a List\<\> of objects.

- When using IQueryable\<T\> as datasource, you can use any public
  property of T in the report. So if type T has a public property
  \"LastName\", you can access it with \<\#dt.LastName\>.

- **Master detail with implicit relationships**. When a public
  property of a collection of objects is other collection of
  objects, the property is considered as a detail of the main
  collection. In this example \"Elements\" is a property of
  \"Categories\", and so there is an implicit relationship between
  them.

- **Master detail with explicit relationships.** While when using LINQ
  you will normally use implicit relationships, you can also relate
  any two collections of objects with a relationship, as you could
  with datasets. In this example, \"ElementName\" is explicitly
  related to \"Elements\" with a call to [FlexCelReport.AddRelationship](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.Report/FlexCelReport/AddRelationship.html).
