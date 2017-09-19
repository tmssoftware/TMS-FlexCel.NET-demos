# LangWars

This example shows how to do a simple report doing FlexCel. It will fetch the most used
tags from Stack Overflow and rank and provide a chart to visualize them. You can work with
online or offline data, in case you don't have access to Stack Overflow.

In this example, we are exporting the results to an html file and showing the results in a web browser.
We could also export to pdf, as shown in the "FlexView" demo.

**IMPORTANT NOTE**  At the time of writing this demo, FlexCel will only render charts in xls files,
not xlsx, so you need to have an xls template and final file to see the chart in the screen.
FlexCel preserves charts in xlsx files, but it won't render them yet. (that is, it won't convert them
to pdf/html/etc).