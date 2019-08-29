using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;

using FlexCel.Core;
using FlexCel.XlsAdapter;
using FlexCel.Render;
using FlexCel.AspNet;

public partial class _Default : System.Web.UI.Page 
{
    protected void Page_Load(object sender, EventArgs e)
    {
        XlsFile xls = new XlsFile();

        string DefaultFile = Server.MapPath("~/default.xls");
        if (IsPostBack)
        {
            if (Uploader.HasFile)
            {
                xls.Open(Uploader.FileContent);
            }
            else
            {
                xls.Open(DefaultFile);
            }
        }
        else
        {
            xls.Open(DefaultFile);
        }

        Viewer.HtmlExport.ImageNaming = TImageNaming.Guid;
        Viewer.HtmlExport.Workbook = xls;
        Viewer.RelativeImagePath = "images";
        Viewer.HtmlExport.FixIE6TransparentPngSupport = true;  //This is only needed if you are using IE and there are transparent png files.
        Viewer.ImageExportMode = TImageExportMode.TemporaryFiles;

    }
}
