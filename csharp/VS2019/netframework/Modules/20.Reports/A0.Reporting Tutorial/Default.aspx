<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Default.aspx.cs" Inherits="_Default" %>

<%@ Register assembly="FlexCel.AspNet" namespace="FlexCel.AspNet" tagprefix="cc1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Untitled Page</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
    </div>
    <asp:Label ID="Label1" runat="server" Text="Welcome to the FlexCel.NET Demo!"></asp:Label>
    <br />
    <asp:Button ID="Button1" runat="server" onclick="Button1_Click" 
        Text="View Report!" />
    <asp:Button ID="Button2" runat="server" onclick="Button2_Click" 
        Text="Export to Xls" />
    <asp:Button ID="Button3" runat="server" onclick="Button3_Click"  
        Text="Export to PDF" />
    <cc1:FlexCelAspViewer ID="FlexCelAspViewer1" runat="server" />
    </form>
</body>
</html>
