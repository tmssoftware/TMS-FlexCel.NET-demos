# FlexCel image explorer

This is an application that will show all the images in an Excel file, to show
their real size and bit depth. It will also allow you to extract those
images to files. Note that the first time you open a file it can take
some time, as it opens **ALL** the xls files on the folder and extracts
all the image information so it can highlight files with problems.

## Concepts

- How to use the API to read the images on a file.

- How to convert the images to 256 colors or black and white using
  GDI+

- On the left pane, files in red are files that have a cropped image.
  (that is, the image stored is larger than the image shown on
  Excel). Having this cropped image only consumes disk space, and
  makes rendering to pdf slower, as the image has to be decoded,
  cropped to the real size and encoded again.

- **Bold** entries on the left pane are images with true color and
  transparency. Those images can get very big quite fast, so it is
  better if you can convert them to indexed color. Also, file with
  true color and transparency have to be re-encoded when exporting to
  pdf, making the process slower.
