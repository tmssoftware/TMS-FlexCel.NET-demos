# Using hyperlinks

This is a small application to show how to use hyperlinks. Parameters in
a hyperlink are not straightforward, so the idea is that you can go to
Excel, create a Worksheet with your desired hyperlinks, and use this app
to see what parameters they have.

## Columns in the Grid

- **Cell1, Cell2**:

  These are the first and last cell on the range of the hyperlink.
  Normally they will be the same cell

- **Type:** There are 4 types of Hyperlinks:

   1. **URL**: This can be http, https, ftp, or mailto://
   2. **UNC**: This is a path to a network site on universal naming
   convention, like \\\\server\\folder\\your\_file.xls.

   3. **Local File**: This is a file stored relative to the path of the
   sheet. (for example, on the upper folder)

   4. **Current Workbook**: this is a link to a cell on the file. Note
   that this option always has Text=\'\'

- **Text:**

  Text of the HyperLink. This is empty when linking to a cell.

- **Description:**

  Description of the HyperLink.

- **Target Frame:**

  This parameter is not documented. You can leave it empty.

- **Text Mark:**

  When entering an URL on Excel, you can enter additional text following
  the url with a \"\#\" character (for example
  www.your\_url.com\#myurl\") The text Mark is the text after the \"\#\"
  char. When entering an address to a cell, the address goes here too.

- **Hint:**

  This is the hint Excel will show when hovering the mouse over the
  hyperlink.
