# Handling non-fatal errors

On this demo we are going to see how to deal with errors that are not
fatal and will normally be ignored, but that can degrade the generated
files.

You can hook your own listener to the [FlexCelTrace](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.Core/FlexCelTrace/index.html) static class to gain
control on what should be done when a non-fatal error happens.

## Concepts

- How to hook a listener to the static [FlexCelTrace](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.Core/FlexCelTrace/index.html) class. In a normal
  application you would just hook the listener at the beginning of
  the application. On this case, since our demo can be used
  standalone or from the MainDemo browser, we need to make sure we
  unhook the event when the form is disposed.

- How to stop the file generation when a non-fatal error happens.

- How we should take care of threading issues, since [FlexCelTrace](https://doc.tmssoftware.com/flexcel/net/api/FlexCel.Core/FlexCelTrace/index.html) is a
  static class that could be called from more than one place.

- How to ignore different types of errors.
