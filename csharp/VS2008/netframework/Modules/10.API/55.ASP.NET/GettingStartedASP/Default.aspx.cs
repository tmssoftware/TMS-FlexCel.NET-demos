using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;

using System.IO;
using System.Reflection;
using System.Drawing;

using FlexCel.Core;
using FlexCel.Render;
using FlexCel.XlsAdapter;

public partial class _Default : System.Web.UI.Page 
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }

    private void CreateFile(ExcelFile Xls)
    {
        //Create a new file. We could also open an existing file with Xls.Open
        Xls.NewFile(1);
        //Set some cell values.
        Xls.SetCellValue(1, 1, "Hello to everybody");
        Xls.SetCellValue(2, 1, 3);
        Xls.SetCellValue(3, 1, 2.1);
        Xls.SetCellValue(4, 1, new TFormula("=Sum(A2,A3)"));

        //Load an image from disk.
        string AssemblyPath = HttpContext.Current.Request.PhysicalApplicationPath;
        using (System.Drawing.Image Img = System.Drawing.Image.FromFile(Path.Combine(Path.Combine(AssemblyPath, "images"), "Test.bmp")))
        {

            //Add a new image on cell E5
            Xls.AddImage(2, 6, Img);
            //Add a new image with custom properties at cell F6
            Xls.AddImage(Img, new TImageProperties(new TClientAnchor(TFlxAnchorType.DontMoveAndDontResize, 2, 10, 6, 10, 100, 100, Xls), ""));
            //Swap the order of the images. it is not really necessary here, we could have loaded them on the inverse order.
            Xls.BringToFront(1);
        }

        //Add a comment on cell a2
        Xls.SetComment(2, 1, "This is 3");

        //Custom Format cells a2 and a3
        TFlxFormat f = Xls.GetDefaultFormat;
        f.Font.Name = "Times New Roman";
        f.Font.Color = Color.Red;
        f.FillPattern.Pattern = TFlxPatternStyle.LightDown;
        f.FillPattern.FgColor = Color.Blue;
        f.FillPattern.BgColor = Color.White;

        int XF = Xls.AddFormat(f);

        Xls.SetCellFormat(2, 1, XF);
        Xls.SetCellFormat(3, 1, XF);

        f.Rotation = 45;
        f.FillPattern.Pattern = TFlxPatternStyle.Solid;
        int XF2 = Xls.AddFormat(f);
        //Apply a custom format to all the row.
        Xls.SetRowFormat(1, XF2);

        //Merge cells
        Xls.MergeCells(5, 1, 10, 6);
        //Note how this one merges with the previous range, creating a final range (5,1,15,6)
        Xls.MergeCells(10, 6, 15, 6);

        //Make sure rows are autofitted for pdf export.
         Xls.AutofitRowsOnWorkbook(false, true, 1);

    }


    protected void BtnReadCellA1_Click(object sender, EventArgs e)
    {
        ExcelFile Xls = new XlsFile();
        if (FileBox.PostedFile == null || FileBox.PostedFile.InputStream == null || FileBox.PostedFile.InputStream.Length == 0)
        {

            LabelA1.Text = "No file selected";
            return;
        }
        FileBox.PostedFile.InputStream.Position = 0;
        try
        {
            Xls.Open(FileBox.PostedFile.InputStream);
            object v = Xls.GetCellValue(1, 1);
            if (v == null) LabelA1.Text = "Cell A1 is empty";
            else LabelA1.Text = "Cell A1 has the value: " + Convert.ToString(v);
        }
        catch (Exception ex)
        {
            LabelA1.Text = ex.Message;
        }

    }
    protected void BtnXls_Click(object sender, EventArgs e)
    {
        ExcelFile Xls = new XlsFile();
        CreateFile(Xls);
        using (MemoryStream ms = new MemoryStream())
        {
            Xls.Save(ms);
            ms.Position = 0;
            Response.Clear();
            Response.AddHeader("Content-Disposition", "attachment; filename=Test.xls");
            Response.AddHeader("Content-Length", ms.Length.ToString());
            Response.ContentType = "application/excel"; //octet-stream";
            Response.BinaryWrite(ms.ToArray());
            Response.End();
        }
    }
    protected void BtnPdf_Click(object sender, EventArgs e)
    {
        ExcelFile Xls = new XlsFile();
        CreateFile(Xls);

        FlexCelPdfExport Pdf = new FlexCelPdfExport(Xls);

        using (MemoryStream ms = new MemoryStream())
        {

            Pdf.BeginExport(ms);
            try
            {
                Pdf.ExportAllVisibleSheets(true, "Getting Started");
            }
            finally
            {
                Pdf.EndExport();
            }
            ms.Position = 0;
            Response.Clear();
            Response.AddHeader("Content-Disposition", "attachment; filename=Test.pdf");
            Response.AddHeader("Content-Length", ms.Length.ToString());
            Response.ContentType = "application/pdf"; //octet-stream";
            Response.BinaryWrite(ms.ToArray());
            Response.End();
        }

    }
}
