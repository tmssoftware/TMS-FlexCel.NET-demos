# Virtual datasets

Normally FlexCelReport uses DataTables for its data. On reporting we can
see DataTables as \"overcharged\" arrays, with added functionality like
filtering or sorting. They are also really powerful and fast, so using
them is normally the best option. But in some cases you might have very
large bussiness objects and would like to use them directly without
copying them first into a DataTable. On those cases, you can create your
own **VirtualDataset** and **VirtualDatasetState** descendants to do
this task.

This is an advanced topic. Remember to read the  [Appendix: Virtual DataSets
 on](https://doc.tmssoftware.com/flexcel/net/guides/reports-developer-guide.html#appendix:-virtual-datasets
-on) the Reports developer guide.

## Concepts

- How to create a VirtualDataset descendant to encapsulate an
  arbitrary object and make a report with it. Two descendants are
  shown, one very simple with the minimum functionality needed, and
  other that fully implements all the features.

- As you can see on the *ComplexVirtualArrayDataSource* example, to
  provide all the functionality you do not need to override too many
  methods. But you should implement very efficient methods. If you
  do not, probably the performance dumping everything to a DataSet
  would be faster.

- How to create \"infographics\". this is based on the article:
  http://www.juiceanalytics.com/weblog/?p=236
