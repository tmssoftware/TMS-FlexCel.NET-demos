# Windows RT and Windows Phone 8.1 Examples

The examples in this section apply to FlexCel for Windows 8.1. 

Windows 8.1 requires an async API, so the FlexCel API is different from the rest (you need to use await xls.OpenAsync instead of xls.Open for example). For most use cases you should use FlexCel for UWP instead of FlexCel for Windows 8.1, since it provides a similar API to the other FlexCel implementations. 

Only use FlexCel for Windows 8.1 
if you need to support Windows 8.1 store apps. If your target is Windows 10 store apps, then FlexCel for UWP is the best choice.
