using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Data.OleDb;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using FlexCel.Report;
using FlexCel.Render;
using FlexCel.Pdf;
using System.Globalization;

using System.Xml;


namespace FeaturesPage
{
    public partial class mainForm: System.Windows.Forms.Form
    {
        public mainForm()
        {
            InitializeComponent();
            //initialize the db.
            dbconnection.ConnectionString = dbconnection.ConnectionString.Replace("Features.mdb", Path.Combine(DataPath, "features.mdb"));
            ResizeToolbar(mainToolbar);
            FlexCelConfig.DpiForImages = 192; //Make the exports in hidpi.
        }

        private void ResizeToolbar(ToolStrip toolbar)
        {

            using (Graphics gr = CreateGraphics())
            {
                double xFactor = gr.DpiX / 96.0;
                double yFactor = gr.DpiY / 96.0;
                toolbar.ImageScalingSize = new Size((int)(24 * xFactor), (int)(24 * yFactor));
                toolbar.Width = 0; //force a recalc of the buttons.
            }
        }

        private void button2_Click(object sender, System.EventArgs e)
        {
            Close();
        }

        private static string DataPath
        {
            get
            {
                return Path.Combine(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), ".."), "..") + Path.DirectorySeparatorChar;
            }
        }

        private string ResultPath
        {
            get
            {
                string BasePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                return Path.Combine(BasePath, "Features");
            }
        }

        private XlsFile Export(DataSet data)
        {
            using (FlexCelReport Report = new FlexCelReport(true))
            {
                Report.AddTable(data);
                Report.SetUserFunction("Images", new ImagesImp());
                XlsFile Xls = new XlsFile(true);
                Xls.Open(Path.Combine(DataPath, "Features Page.template.xls"));

                Report.Run(Xls);
                return Xls;
            }

        }

        private DataSet LoadDataSet()
        {
            DataSet Result = new DataSet();
            featuresAdapter.Fill(Result, "Features");
            categoriesAdapter.Fill(Result, "Categories");
            hyperlinksAdapter.Fill(Result, "Hyperlinks");
            Result.Relations.Add(Result.Tables["Categories"].Columns["CategoryId"], Result.Tables["Features"].Columns["CategoryId"]);
            Result.Relations.Add(Result.Tables["Features"].Columns["FeaturesId"], Result.Tables["Hyperlinks"].Columns["FeaturesId"]);

            return Result;
        }

        private void btnExportExcel_Click(object sender, System.EventArgs e)
        {

            string XlsPath = Path.Combine(ResultPath, "FeaturesFlexCel.xls");
            using (DataSet data = LoadDataSet())
            {
                XlsFile Xls = Export(data);

                Directory.CreateDirectory(ResultPath);
                Xls.Save(XlsPath);
            }

            if (MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                Process.Start(XlsPath);
            }

        }

        private void btnExportHtml_Click(object sender, System.EventArgs e)
        {
            string MainHtmlPath = Path.Combine(ResultPath, "featuresflexcel.htm");

            using (DataSet data = LoadDataSet())
            {
                XlsFile Xls = Export(data);

                Directory.CreateDirectory(ResultPath);
                using (FlexCelHtmlExport html = new FlexCelHtmlExport(Xls, true))
                {
                    html.ImageResolution = 192;
                    html.ImageBackground = Color.White; //Since we are not setting html.FixIE6TransparentPngSupport, we must ensure tehre are no transparent images.
                    TStandardSheetSelector SheetSelector = new TStandardSheetSelector(TSheetSelectorPosition.Top);
                    SheetSelector.SheetSelectorEntry += new SheetSelectorEntryEventHandler(SheetSelector_SheetSelectorEntry);
                    SheetSelector.CssGeneral.Main += "font-family:Verdana;font-size:10pt;";

                    html.ExportAllVisibleSheetsAsTabs(ResultPath, "Features", ".htm", null, null, SheetSelector);

                    //Rename the first tab so it is "featuresflexcel.htm";
                    string[] Sheets = html.GeneratedFiles.GetHtmlFiles();
                    File.Delete(MainHtmlPath);
                    File.Move(Sheets[0], MainHtmlPath);

                }
            }
            if (MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                Process.Start(MainHtmlPath);
            }


        }

        static void SheetSelector_SheetSelectorEntry(object sender, SheetSelectorEntryEventArgs e)
        {
            //We will rename the first sheet, so we need to update the links here.
            if (e.ActiveSheet == 1) e.Link = "featuresflexcel.htm";
        }

        private void btnExportPDF_Click(object sender, System.EventArgs e)
        {
            string PdfPath = Path.Combine(ResultPath, "FeaturesFlexCel.pdf");

            using (DataSet data = LoadDataSet())
            {
                XlsFile Xls = Export(data);
                Directory.CreateDirectory(ResultPath);

                using (FlexCelPdfExport pdf = new FlexCelPdfExport(Xls, true))
                {
                    using (FileStream pdfStream = new FileStream(PdfPath, FileMode.Create))
                    {
                        pdf.BeginExport(pdfStream);
                        pdf.FontMapping = TFontMapping.ReplaceAllFonts;

                        pdf.Properties.Subject = "A list of FlexCel.NET features";
                        pdf.Properties.Author = "TMS Software";
                        pdf.Properties.Title = "List of FlexCel.NET features";
                        pdf.PageLayout = TPageLayout.Outlines;
                        pdf.ExportAllVisibleSheets(false, "Features");
                        pdf.EndExport();
                    }
                }
            }

            if (MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                Process.Start(PdfPath);
            }

        }


        class ImagesImp: TFlexCelUserFunction
        {
            public ImagesImp()
            {
            }

            public override object Evaluate(object[] parameters)
            {
                if (parameters == null || parameters.Length != 1)
                    throw new ArgumentException("Bad parameter count in call to Images() user-defined function");

                string ImageFilename = Path.Combine(Path.Combine(DataPath, "images"), "Features" + Convert.ToString(parameters[0], CultureInfo.InvariantCulture) + ".png");
                if (File.Exists(ImageFilename))
                {
                    using (FileStream fs = new FileStream(ImageFilename, FileMode.Open))
                    {
                        byte[] Result = new byte[fs.Length];
                        fs.Read(Result, 0, Result.Length);
                        return Result;
                    }
                }

                return null;
            }
        }

    }

}
