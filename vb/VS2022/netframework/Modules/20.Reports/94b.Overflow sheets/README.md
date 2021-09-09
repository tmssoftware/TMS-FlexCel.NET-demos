# Overflowing data in different sheets.

**THIS DEMO NEEDS .NET 3.5 OR NEWER. YOU COULD DO THE SAME IN .NET 2.0,
BUT THIS PARTICULAR DEMO USES LINQ.**

For very big datasets, it might happen that they don\'t fit in a single
sheet. In xls files, you can have a maximum of 65536 rows, and in xlsx
1048576 rows. If you want to dump bigger datasets than the maximum
available (or you just want to allow say a maximum of 100 records in a
sheet and after it is reached move to the next), then you might use the
SPLIT tag as a master in the sheet.

Take a look at [Splitting Tables](https://doc.tmssoftware.com/flexcel/net/guides/reports-designer-guide.html#splitting-tables)
in the Report designer's guide for more information.

## Concepts

- How to use the SPLIT tag to create a master-detail relationship of a
  dataset with itself, where each detail has a fixed number of
  records. This way you can limit the number of records that you
  show in each sheet. 
  
- How to use the [Evaluate](https://doc.tmssoftware.com/flexcel/net/guides/reports-tag-reference.html#evaluate) tag to provide a meaningful sheet name.
