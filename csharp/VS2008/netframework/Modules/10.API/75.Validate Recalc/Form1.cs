using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using System.IO;
using System.Diagnostics;
using System.Reflection;

namespace ValidateRecalc
{
    /// <summary>
    /// Use this demo to validate the recalculation mad by FlexCel.
    /// </summary>
    public partial class mainForm: System.Windows.Forms.Form
    {
        private FlexCel.Report.FlexCelReport XlsReport;

        public mainForm()
        {
            InitializeComponent();
            ResizeToolbar(mainToolbar);
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

        private void btnInfo_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This example will validate the calculations performed by the FlexCel engine.\n" +
            "It can do it in 2 different ways:\n" +
            "  1) The button 'Validate Recalc' will analyze a file, and report if there is anything that FlexCel doesn't support on it.\n" +
            "  2) The button 'Compare with Excel' will open a file saved by Excel, recalculate it with FlexCel, compare the values reported by both FlexCel and Excel and report if there are any differences.");
        }

        private void validateRecalc_Click(object sender, System.EventArgs e)
        {
            if (openFileDialog1.ShowDialog() != DialogResult.OK) return;
            XlsFile Xls = new XlsFile();

            Xls.Open(openFileDialog1.FileName);

            // /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // ////////Code here is only needed if you have linked files. In this example we don't know, so we will use it /////////
            TWorkspace Work = new TWorkspace(); //Create a workspace
            Work.Add(Path.GetFileName(openFileDialog1.FileName), Xls);  //Add the original file to it
            Work.LoadLinkedFile += new LoadLinkedFileEventHandler(Work_LoadLinkedFile);  //Set up an event to load the linked files.
                                                                                         // /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            report.Text = "Results on file: " + openFileDialog1.FileName;
            TUnsupportedFormulaList Usl = Xls.RecalcAndVerify();
            if (Usl.Count == 0)
            {
                report.Text += "\n**********All formulas supported!**********";
                return;
            }

            report.Text += "\nIssues Found:";
            for (int i = 0; i < Usl.Count; i++)
            {
                string FileName = String.Empty;
                if (Usl[i].FileName != null) FileName = "File: " + Usl[i].FileName + "  => ";
                report.Text += "\n     " + FileName + Usl[i].Cell.CellRef + ": " + Usl[i].ErrorType.ToString();
                if (Usl[i].FunctionName != null)
                {
                    string FunctionStr = "Function";
                    if (Usl[i].ErrorType == TUnsupportedFormulaErrorType.ExternalReference) FunctionStr = "Linked file not found";
                    report.Text += " ->" + FunctionStr + ": " + Usl[i].FunctionName;
                }
            }
        }

        private void compareWithExcel_Click(object sender, System.EventArgs e)
        {
            if (openFileDialog1.ShowDialog() != DialogResult.OK) return;
            compareWithExcel.Enabled = false;
            validateRecalc.Enabled = false;
            try
            {
                XlsFile xls1 = new XlsFile();
                XlsFile xls2 = new XlsFile();

                xls1.Open(openFileDialog1.FileName);
                xls2.Open(openFileDialog1.FileName);
                report.Text = "Compare with Excel: " + openFileDialog1.FileName;

                // /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                // ////////Code here is only needed if you have linked files. In this example we don't know, so we will use it /////////
                TWorkspace Work = new TWorkspace(); //Create a workspace
                Work.Add(Path.GetFileName(openFileDialog1.FileName), xls1);  //Add the original file to it
                Work.LoadLinkedFile += new LoadLinkedFileEventHandler(Work_LoadLinkedFile);  //Set up an event to load the linked files.
                                                                                             // /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


                CompareXls(xls1, xls2, null);
            }
            finally
            {
                compareWithExcel.Enabled = true;
                validateRecalc.Enabled = true;
            }
        }

        private void CompareXls(XlsFile xls1, XlsFile xls2, DataTable table)
        {
            int DiffCount = 0;
            xls1.Recalc();

            for (int sheet = 1; sheet <= xls1.SheetCount; sheet++)
            {
                xls1.ActiveSheet = sheet;
                xls2.ActiveSheet = sheet;
                int aColCount = xls1.ColCount;
                for (int r = 1; r <= xls1.RowCount; r++)
                    for (int c = 1; c <= aColCount; c++)
                    {
                        TFormula f = xls1.GetCellValue(r, c) as TFormula;
                        if (f != null)
                        {
                            TCellAddress ad = new TCellAddress(r, c);
                            TFormula f2 = (TFormula)xls2.GetCellValue(r, c);
                            if (f.Result == null) f.Result = "";
                            if (f2.Result == null) f2.Result = "";
                            double eps = 0;
                            if (f.Result is Double && f2.Result is Double)
                            {
                                if ((Double)f2.Result == 0)
                                {
                                    if (Math.Abs((double)f.Result) < Double.Epsilon)
                                        eps = 0;
                                    else
                                        eps = Double.NaN;
                                }
                                else
                                    eps = (double)f.Result / (Double)f2.Result;
                                if (Math.Abs(eps - 1) < 0.001)
                                    f.Result = f2.Result;
                            }
                            if (!f.Result.Equals(f2.Result))
                            {
                                if (table == null)
                                {
                                    report.Text += "\nSheet:" + xls1.SheetName + " --- Cell:" + ad.CellRef + " --- Calculated: " + f.Result.ToString() + "    Excel: " + f2.Result.ToString() + "  dif: " + eps.ToString() + "   formula: " + f.Text;
                                    Application.DoEvents();
                                }
                                else
                                {
                                    table.Rows.Add(new object[] { xls1.SheetName, ad.CellRef, f.Result.ToString(), f2.Result.ToString(), eps.ToString(), f.Text });
                                }
                                DiffCount++;

                            }
                        }
                    }
            }

            if (table == null)
            {
                report.Text += "\nFinished Comparing.";
                if (DiffCount == 0) report.Text += "\n**********No differences found!**********";
                else
                    report.Text += String.Format("\n  --->Found {0} differences", DiffCount);
            }
        }

        private void ValidateXls(XlsFile xls, DataTable table)
        {
            TUnsupportedFormulaList Usl = xls.RecalcAndVerify();
            for (int i = 0; i < Usl.Count; i++)
            {
                table.Rows.Add(new object[]
                    {
                        Usl[i].FileName,
                        Usl[i].Cell.CellRef,
                        Usl[i].ErrorType.ToString(),
                        Usl[i].FunctionName
                    });
            }
        }

        /// <summary>
        /// This is the method that will be called by the ASP.NET front end. It returns an array of bytes 
        /// with the report data, so the ASP.NET application can stream it to the client.
        /// </summary>
        /// <returns>The generated file as a byte array.</returns>
        public byte[] WebRun(Stream DataStream, string FileName)
        {
            XlsReport.SetValue("Date", DateTime.Now);
            XlsReport.SetValue("FileName", FileName);
            DataSet Data = new DataSet();
            DataTable ValidateResult = Data.Tables.Add("ValidateResult");
            ValidateResult.Columns.Add("FileName", typeof(string));
            ValidateResult.Columns.Add("CellRef", typeof(string));
            ValidateResult.Columns.Add("ErrorType", typeof(string));
            ValidateResult.Columns.Add("FunctionName", typeof(string));

            DataTable CompareResult = Data.Tables.Add("CompareResult");
            CompareResult.Columns.Add("SheetName", typeof(string));
            CompareResult.Columns.Add("CellRef", typeof(string));
            CompareResult.Columns.Add("CalcResult", typeof(string));
            CompareResult.Columns.Add("XlsResult", typeof(string));
            CompareResult.Columns.Add("Diff", typeof(string));
            CompareResult.Columns.Add("FormulaText", typeof(string));

            XlsReport.AddTable(Data);

            XlsFile xls1 = new XlsFile();
            XlsFile xls2 = new XlsFile();

            xls1.Open(DataStream);
            DataStream.Position = 0;
            xls2.Open(DataStream);

            CompareXls(xls1, xls2, CompareResult);
            ValidateXls(xls1, ValidateResult);

            string DataPath = Path.Combine(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), ".."), "..") + Path.DirectorySeparatorChar;

            using (MemoryStream OutStream = new MemoryStream())
            {
                using (FileStream InStream = new FileStream(DataPath + "ValidateReport.xls", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    XlsReport.Run(InStream, OutStream);
                    return OutStream.ToArray();
                }
            }
        }

        /// <summary>
        /// This event is used when there are linked files, to load them on demand.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Work_LoadLinkedFile(object sender, LoadLinkedFileEventArgs e)
        {
            //IMPORTANT: DO NOT USE THIS METHOD IN PRODUCTION IF SECURITY IS IMPORTANT.
            //This method will access any file in your harddisk, as long as it is linked in the spreaadhseet, and
            //that could mean an IMPORTANT SECURITY RISK. You should limit the places where the app can search for 
            //linked files. Look at the "Recalculating Linked Files" in the PDF API Guide for more information.

            string FilePath = Path.Combine(Path.GetDirectoryName(openFileDialog1.FileName), e.FileName);

            if (File.Exists(FilePath)) //If we find the path, just load the file.
            {
                e.Xls = new XlsFile();
                e.Xls.Open(FilePath);
                return;
            }

            //If we couldn't find the file, ask the user for its location.
            linkedFileDialog.FileName = FilePath;
            if (linkedFileDialog.ShowDialog() != DialogResult.OK) return;  //if user cancels, e.Xls will be null, so no file will be used and an #errna error will show in the formulas.

            e.Xls = new XlsFile();
            e.Xls.Open(linkedFileDialog.FileName);

        }

    }

}

