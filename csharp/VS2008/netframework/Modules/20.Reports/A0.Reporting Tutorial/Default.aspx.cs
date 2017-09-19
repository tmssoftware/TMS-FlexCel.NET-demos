using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using FlexCel.Render;
using FlexCel.Report;
using System.IO;

public partial class _Default : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }

    private ExcelFile CreateReport()
    {
        XlsFile Result = new XlsFile(true);
        Result.Open(MapPath("~/App_Data/template.xls"));
        using (FlexCelReport fr = new FlexCelReport())
        {
            LoadData(fr);

            fr.SetValue("ReportCaption", "Hello from FlexCel!");
            fr.Run(Result);
            return Result;
        }
    }

    private void LoadData(FlexCelReport fr)
    {
        DataSet1 Data = new DataSet1();
        DataSet1TableAdapters.ProductTableAdapter ProductAdapter = new DataSet1TableAdapters.ProductTableAdapter();
        ProductAdapter.Fill(Data.Product);

        DataSet1TableAdapters.ProductPhotoTableAdapter ProductPhotoAdapter = new DataSet1TableAdapters.ProductPhotoTableAdapter();
        ProductPhotoAdapter.Fill(Data.ProductPhoto);

        DataSet1TableAdapters.ProductProductPhotoTableAdapter ProductProductPhotoAdapter = new DataSet1TableAdapters.ProductProductPhotoTableAdapter();
        ProductProductPhotoAdapter.Fill(Data.ProductProductPhoto);

        fr.AddTable(Data);
    }

    protected void Button1_Click(object sender, EventArgs e)
    {
        ExcelFile xls = CreateReport();
        FlexCelAspViewer1.HtmlExport.Workbook = xls;
    }
    protected void Button2_Click(object sender, EventArgs e)
    {
        ExcelFile xls = CreateReport();

        using (MemoryStream ms = new MemoryStream())
        {
            xls.Save(ms);
            ms.Position = 0;
            Response.Clear();
            Response.AddHeader("Content-Disposition", "attachment; filename=Test.xls");
            Response.AddHeader("Content-Length", ms.Length.ToString());
            Response.ContentType = "application/excel"; //octet-stream";
            Response.BinaryWrite(ms.ToArray());
            Response.End();
        }


    }
    protected void Button3_Click(object sender, EventArgs e)
    {
        ExcelFile xls = CreateReport();

        using (MemoryStream ms = new MemoryStream())
        {
            using (FlexCelPdfExport pdf = new FlexCelPdfExport())
            {
                pdf.Workbook = xls;
                pdf.BeginExport(ms);
                pdf.ExportAllVisibleSheets(false, "FlexCel");
                pdf.EndExport();
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
}

