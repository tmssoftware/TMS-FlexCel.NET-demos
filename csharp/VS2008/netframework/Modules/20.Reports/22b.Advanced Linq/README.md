# Advanced Linq

**THIS DEMO REQUIRES .NET 3.5 OR NEWER.**

Most of the demos here use datasets as datasources. This is just for
convenience, because Linq is not supported in .NET 2.0 and so if we used
Linq those demos wouldn\'t work for everybody, and also because the
focus is in the Excel templates, not so much in the data layer. But you
can use any IQueryable\<T\> collection as a datasource in a FlexCel
report, and this is what we will show here.

This demo shows some features not shown in the [Linq](https://download.tmssoftware.com/flexcel/doc/net/samples/csharp/netframework/reports/linq/index.html) example.

## Concepts

- How to do a master-detail report when the details are nested many
  levels inside the master. In this case, the class **Country** has
  a **People** class, and the People class has a list of
  **Language** objects. If People was a List\<\> inside Country and
  you wanted to use that list, you would just define a
  **\_\_People\_\_** band (this is shown in the [Linq](https://download.tmssoftware.com/flexcel/doc/net/samples/csharp/netframework/reports/linq/index.html) example). But
  as the List\<\> is inside People which in turn is inside Country,
  you need to define a **\_\_People.Language\_\_** band.

- How to reference a table with dots using **\[square brackets\]**. If
  you write in a cell \<\#tablename.**section.field**\> FlexCel will
  interpret this as table "tablename", field "section.field". The
  text up to the first dot is always the table, and the rest is the
  field. But sometimes you might want this to being interpreted as
  table "tablename.section", field "field". To do so, you need to
  write \<\#**\[tablename.section\]**.field\>. In this particular
  case, we have a table \"People.Language\" which we defined in the
  previous point. If we wrote in cell B1:
  \"\<\#people.language.speakers.percent\> FlexCel would interpret
  this is the table \"people\", not \"people.language\" which is
  what we need. To make FlexCel understand that we want a table
  \"people.language\" we use
  **\<\#\[people.language\].speakers.percent\>**
