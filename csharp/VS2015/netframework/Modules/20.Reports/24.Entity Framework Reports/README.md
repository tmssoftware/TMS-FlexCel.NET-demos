# Using Entity Framework datasources

> [!Note]
> 
> To run this example, you need to have SQL Server (LocalDB) installed.
> For this reason, and because it uses NuGet, this demo doesn't run from
> MainDemo. You need to open this solution and run it alone.


While you can use any generic IQueryable\<T\> object as a source of
reports, here we provide a concrete example of redoing the \"range
reports\" demo using Entity Framework.

## Concepts

- Entity Framework reports should be run inside a **serializable** or
  **snapshot** transaction. While snapshot transactions are
  preferred as they won\'t block other users from changing the data
  while the report is running, in this example we use serializable
  (Because LocalDB doesn\'t support snapshot)

- Here we have an implicit relationship between categories and
  products, so we don't need to add an explicit one. The report
  will just work.
