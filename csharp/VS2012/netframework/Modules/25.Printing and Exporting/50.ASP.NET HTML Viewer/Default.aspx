<%@ Page Language="C#" AutoEventWireup="true"  CodeFile="Default.aspx.cs" Inherits="_Default" %>

<%@ Register Assembly="FlexCel.AspNet" Namespace="FlexCel.AspNet" TagPrefix="cc2" %>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>FlexCel ASP.NET Viewer Demo</title>
</head>
<body id="flexcel1">
    <form id="form1" runat="server">
        A small demo to show Excel files inside a browser.<br />
        Please select a file to view:<br />
        <asp:FileUpload ID="Uploader" runat="server" /><br />
        <asp:Button ID="Button1" runat="server" Text="Load" Width="106px" /><br />
        <hr />
        &nbsp;<cc2:flexcelaspviewer id="Viewer" runat="server" SheetExport="AllVisibleSheets" ImageExportMode="TemporaryFiles"></cc2:flexcelaspviewer>
        
    </form>
</body>
</html>
