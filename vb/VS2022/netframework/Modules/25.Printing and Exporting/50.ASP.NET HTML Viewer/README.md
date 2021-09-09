# Using the FlexCel ASP viewer component

A simple example on how to use the component FlexCelAspViewer to display
xls files as html in a web browser.

> [!Note]
> This demo **does not run from MainDemo** and it needs **.NET 2.0
> or newer.


In order to run this demo, open the website from Visual Studio 2005 or
newer, by doing: Menu-\>File-\>Open-\>Website

and pointing to the Modules\\25.Printing and Exporting\\50.ASP.NET HTML
Viewer folder.

## Concepts

- In this demo we will use .TemporaryFiles to create
  the images. This will create all images as GUIDs in the \"images\"
  folder. Images older than 15 minutes will be deleted when a new
  page is requested, and there is an \"overflow guard\" that
  prevents having more than 5000 images in the images folder.

- If you wanted to use HttpHandlers to return the images, you would
  need to add the following lines to **web.config:**

  ```xml
  <httpHandlers>
  <add verb="*" path="flexcelviewer.ashx"
    type="FlexCel.AspNet.UniqueTemporaryFilesImageHandler,FlexCel.AspNet"/>
  </httpHandlers>
  ```

  This has already been done in the example, so you just need to switch
  the image mode to use UniqueTemporaryFiles.

- In order to delete the temporary images once the timeout has
  happened, we use ImageTimeout and MaxTemporaryImages. Make
  sure you have rights to delete images in the image folders for
  this to work. You could also manually delete the images with a
  script on the server.

- FlexCelAspViewer is great for small sites, but due to the way it
  handles images, (make sure you read the [
    >) it might not scale enough if you do not provide
  a custom image handler.
](https://doc.tmssoftware.com/flexcel/net/guides/html-exporting-guide.html#
-----it-might-not-scale-enough-if-you-do-not-provide
--a-custom-image-handler
)
