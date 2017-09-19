<%@ Page Language="C#" AutoEventWireup="true"  CodeFile="Default.aspx.cs" Inherits="_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>FlexCel ASP.NET 20 Demo</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:Image ID="Image1" runat="server" AlternateText="FlexCel ASP.NET Demo" ImageUrl="~/images/studio.png" />&nbsp;<br />
        <br />
        <asp:Label ID="Label1" runat="server" Text="Upload a file to the server, and read the contents of cell A1:"></asp:Label>&nbsp;<br />
        <asp:FileUpload ID="FileBox" runat="server" />
        <br />
        <asp:Button ID="BtnReadCellA1" runat="server" OnClick="BtnReadCellA1_Click" Text="Read Cell A1" /><br />
        <asp:Label ID="LabelA1" runat="server" BackColor="#FFFFC0" ForeColor="Red"></asp:Label><br />
        <br />
        <hr />
    
    </div>
        <br />
        <asp:Label ID="Label2" runat="server" Text="Create a new Xls or Pdf file and stream it to the client:"></asp:Label><br />
        <br />
        <asp:Button ID="BtnXls" runat="server" OnClick="BtnXls_Click" Text="Create Xls File" />
        <asp:Button ID="BtnPdf" runat="server" OnClick="BtnPdf_Click" Text="Create PDF file" />
    </form>
</body>
</html>
