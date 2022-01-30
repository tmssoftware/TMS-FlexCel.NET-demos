# Charts With Dynamic Series

Examples of how to create a chart with one series per row. **Charts with one series per row are not a common thing to do.** You will normally want your series in columns, as shown in the example [Charts](https://doc.tmssoftware.com/flexcel/net/samples/csharp/netframework/reports/charts/index.html). But if you need to do it, you can use this example as a base.

## Concepts

- How to use [swap series](https://doc.tmssoftware.com/flexcel/net/guides/reports-tag-reference.html#swap-series)\> to create a chart with a series per row, as explained in [Creating charts with dynamic series](xref:ReportsDesignerGuide#creating-charts-with-dynamic-series).

- To use <#swap series> in an embedded chart, you name the chart with a name containing <#swap series>.The tag will be removed from the final chart name. To use <#swap series> in a chart sheet, you write it on the sheet name. And again, the tag will be removed from the final sheet name.

- How to use the [FlexCelReport.CustomizeChart](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.Report/FlexCelReport/CustomizeChart.html) event to further customize the chart.
