using System;
using System.Drawing;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Drawing.Drawing2D;

using System.Diagnostics;
using System.Threading;

using FlexCel.Core;
using FlexCel.XlsAdapter;
using FlexCel.Winforms;
using FlexCel.Render;
using FlexCel.Pdf;

namespace CustomPreview
{
    /// <summary>
    /// Previewer of files.
    /// </summary>
    public partial class mainForm: System.Windows.Forms.Form
    {
        public mainForm() : this(new string[0])
        {
        }

        public mainForm(string[] Args)
        {
            InitializeComponent();
            ResizeToolbar(mainToolbar);
            if (Args.Length > 0)
            {
                LoadFile(Args[0]);
            }

            if (ExcelFile.SupportsXlsx)
            {
                this.openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|Excel 97/2003|*.xls|Excel 2007|*.xlsx;*.xlsm|All files|*.*";
            }

            MainPreview.CenteredPreview = true;
            thumbs.CenteredPreview = true;
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


        private void UpdatePages()
        {
            edPage.Text = String.Format("{0} of {1}", MainPreview.StartPage, MainPreview.TotalPages);
        }

        private void flexCelPreview1_StartPageChanged(object sender, System.EventArgs e)
        {
            UpdatePages();
        }

        private void ChangePages()
        {
            string s = edPage.Text.Trim();
            int pos = 0;
            while (pos < s.Length && s[pos] >= '0' && s[pos] <= '9') pos++;
            if (pos > 0)
            {
                int page = MainPreview.StartPage;
                try
                {
                    page = Convert.ToInt32(s.Substring(0, pos));
                }
                catch (Exception)
                {
                }

                MainPreview.StartPage = page;
            }
            UpdatePages();
        }

        private void edPage_Leave(object sender, System.EventArgs e)
        {
            ChangePages();
        }

        private void edPage_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
                ChangePages();
            if (e.KeyChar == (char)27)
                UpdatePages();
        }

        private void flexCelPreview1_ZoomChanged(object sender, System.EventArgs e)
        {
            UpdateZoom();
        }

        private void UpdateZoom()
        {
            edZoom.Text = String.Format("{0}%", (int)Math.Round(MainPreview.Zoom * 100));
            if (MainPreview.AutofitPreview == TAutofitPreview.None) UpdateAutofitText();
        }

        private void ChangeZoom()
        {
            string s = edZoom.Text.Trim();
            int pos = 0;
            while (pos < s.Length && s[pos] >= '0' && s[pos] <= '9') pos++;
            if (pos > 0)
            {
                int zoom = (int)Math.Round(MainPreview.Zoom * 100);
                try
                {
                    zoom = Convert.ToInt32(s.Substring(0, pos));
                }
                catch (Exception)
                {
                }

                MainPreview.Zoom = zoom / 100.0;
            }
            UpdateZoom();
        }

        private void edZoom_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
                ChangeZoom();
            if (e.KeyChar == (char)27)
                UpdateZoom();
        }

        private void edZoom_Enter(object sender, System.EventArgs e)
        {
            ChangeZoom();
        }

        private void btnClose_Click(object sender, System.EventArgs e)
        {
            Close();
        }

        private void openFile_Click(object sender, System.EventArgs e)
        {
            if (openFileDialog.ShowDialog() != DialogResult.OK) return;
            LoadFile(openFileDialog.FileName);
        }

        //The event that will actually provide the password to open the empty form.
        private void GetPassword(OnPasswordEventArgs e)
        {
            PasswordForm Pwd = new PasswordForm();
            e.Password = string.Empty;
            if (Pwd.ShowDialog() != DialogResult.OK) return;
            e.Password = Pwd.Password;
        }


        internal void LoadFile(string FileName)
        {
            openFileDialog.FileName = FileName;
            lbSheets.Items.Clear();

            XlsFile xls = new XlsFile();
            xls.HeadingColWidth = -1;
            xls.HeadingRowHeight = -1;
            xls.Protection.OnPassword += new OnPasswordEventHandler(GetPassword);
            xls.Open(FileName);

            for (int i = 1; i <= xls.SheetCount; i++)
            {
                lbSheets.Items.Add(xls.GetSheetName(i));
            }

            lbSheets.SelectedIndex = xls.ActiveSheet - 1;

            flexCelImgExport1.Workbook = xls;
            MainPreview.InvalidatePreview();
            Text = "Custom Preview: " + openFileDialog.FileName;
            //btnHeadings.Checked = flexCelImgExport1.Workbook.PrintHeadings;
            //btnGridLines.Checked = flexCelImgExport1.Workbook.PrintGridLines;
            btnFirst.Enabled = true; btnPrev.Enabled = true; btnNext.Enabled = true; btnLast.Enabled = true; edPage.Enabled = true;
            btnZoomIn.Enabled = true; edZoom.Enabled = true; btnZoomOut.Enabled = true;
            btnGridLines.Enabled = true; btnHeadings.Enabled = true; btnRecalc.Enabled = true; btnPdf.Enabled = true;

        }

        private void btnFirst_Click(object sender, System.EventArgs e)
        {
            MainPreview.StartPage = 1;
        }

        private void btnPrev_Click(object sender, System.EventArgs e)
        {
            MainPreview.StartPage--;
        }

        private void btnNext_Click(object sender, System.EventArgs e)
        {
            MainPreview.StartPage++;
        }

        private void btnLast_Click(object sender, System.EventArgs e)
        {
            MainPreview.StartPage = MainPreview.TotalPages;
        }

        private void btnZoomOut_Click(object sender, System.EventArgs e)
        {
            MainPreview.Zoom -= 0.1;
        }

        private void btnZoomIn_Click(object sender, System.EventArgs e)
        {
            MainPreview.Zoom += 0.1;
        }

        private void lbSheets_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            if (flexCelImgExport1.Workbook == null) return;
            if (lbSheets.Items.Count > flexCelImgExport1.Workbook.SheetCount) return;
            flexCelImgExport1.Workbook.ActiveSheet = lbSheets.SelectedIndex + 1;
            MainPreview.InvalidatePreview();
        }

        private void btnPdf_Click(object sender, System.EventArgs e)
        {
            if (flexCelImgExport1.Workbook == null)
            {
                MessageBox.Show("There is no open file");
                return;
            }
            if (PdfSaveFileDialog.ShowDialog() != DialogResult.OK) return;

            using (FlexCelPdfExport PdfExport = new FlexCelPdfExport(flexCelImgExport1.Workbook, true))
            {
                if (!DoExportToPdf(PdfExport)) return;
            }

            if (MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            Process.Start(PdfSaveFileDialog.FileName);
        }

        private bool DoExportToPdf(FlexCelPdfExport PdfExport)
        {
            PdfThread MyPdfThread = new PdfThread(PdfExport, PdfSaveFileDialog.FileName, cbAllSheets.Checked);
            Thread PdfExportThread = new Thread(new ThreadStart(MyPdfThread.ExportToPdf));
            PdfExportThread.Start();
            using (PdfProgressDialog Pg = new PdfProgressDialog())
            {
                Pg.ShowProgress(PdfExportThread, PdfExport);
                if (Pg.DialogResult != DialogResult.OK)
                {
                    PdfExport.Cancel();
                    PdfExportThread.Join(); //We could just leave the thread running until it dies, but there are 2 reasons for waiting until it finishes:
                                            //1) We could dispose it before it ends. This is workaroundable.
                                            //2) We might change its workbook object before it ends (by loading other file). This will surely bring issues.
                    return false;
                }

                if (MyPdfThread != null && MyPdfThread.MainException != null)
                {
                    throw MyPdfThread.MainException;
                }
            }
            return true;
        }

        private void cbAllSheets_CheckedChanged(object sender, System.EventArgs e)
        {
            lbSheets.Visible = !cbAllSheets.Checked;
            sheetSplitter.Visible = lbSheets.Visible;
            flexCelImgExport1.AllVisibleSheets = cbAllSheets.Checked;
            if (flexCelImgExport1.Workbook == null) return;
            MainPreview.InvalidatePreview();

        }

        private void btnRecalc_Click(object sender, System.EventArgs e)
        {
            if (flexCelImgExport1.Workbook == null)
            {
                MessageBox.Show("Please open a file before recalculating.");
                return;
            }
            flexCelImgExport1.Workbook.Recalc(true);
            MainPreview.InvalidatePreview();

        }


        private void mainForm_Load(object sender, System.EventArgs e)
        {
        }

        private void btnHeadings_Click(object sender, EventArgs e)
        {
            ExcelFile xls = flexCelImgExport1.Workbook;
            if (xls == null)
            {
                return;
            }

            if (cbAllSheets.Checked)
            {
                int SaveActiveSheet = xls.ActiveSheet;
                for (int sheet = 1; sheet <= xls.SheetCount; sheet++)
                {
                    xls.ActiveSheet = sheet;
                    xls.PrintHeadings = btnHeadings.Checked;
                }
                xls.ActiveSheet = SaveActiveSheet;
            }
            else
            {
                xls.PrintHeadings = btnHeadings.Checked;
            }
            MainPreview.InvalidatePreview();

        }

        private void btnGridLines_Click(object sender, EventArgs e)
        {
            ExcelFile xls = flexCelImgExport1.Workbook;
            if (xls == null)
            {
                return;
            }

            if (cbAllSheets.Checked)
            {
                int SaveActiveSheet = xls.ActiveSheet;
                for (int sheet = 1; sheet <= xls.SheetCount; sheet++)
                {
                    xls.ActiveSheet = sheet;
                    xls.PrintGridLines = btnGridLines.Checked;
                }
                xls.ActiveSheet = SaveActiveSheet;
            }
            else
            {
                xls.PrintGridLines = btnGridLines.Checked;
            }
            MainPreview.InvalidatePreview();

        }

        private void noneToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MainPreview.AutofitPreview = TAutofitPreview.None;
            UpdateAutofitText();
        }

        private void fitToWidthToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MainPreview.AutofitPreview = TAutofitPreview.Width;
            UpdateAutofitText();
        }

        private void fitToHeightToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MainPreview.AutofitPreview = TAutofitPreview.Height;
            UpdateAutofitText();
        }

        private void fitToPageToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MainPreview.AutofitPreview = TAutofitPreview.Full;
            UpdateAutofitText();
        }

        private void UpdateAutofitText()
        {
            switch (MainPreview.AutofitPreview)
            {
                case TAutofitPreview.None:
                    btnAutofit.Text = "No Autofit";
                    break;
                case TAutofitPreview.Width:
                    btnAutofit.Text = "Fit to Width";
                    break;
                case TAutofitPreview.Height:
                    btnAutofit.Text = "Fit to Height";
                    break;
                case TAutofitPreview.Full:
                    btnAutofit.Text = "Fit to Page";
                    break;
                default:
                    break;
            }

        }

    }

    #region PdfThread
    class PdfThread
    {
        private FlexCelPdfExport PdfExport;
        private string FileName;
        private bool AllVisibleSheets;
        private Exception FMainException;

        internal PdfThread(FlexCelPdfExport aPdfExport, string aFileName, bool aAllVisibleSheets)
        {
            PdfExport = aPdfExport;
            FileName = aFileName;
            AllVisibleSheets = aAllVisibleSheets;
        }

        internal void ExportToPdf()
        {
            try
            {
                if (AllVisibleSheets)
                {
                    try
                    {
                        using (FileStream f = new FileStream(FileName, FileMode.Create, FileAccess.Write))
                        {
                            PdfExport.BeginExport(f);
                            PdfExport.PageLayout = TPageLayout.Outlines;
                            PdfExport.ExportAllVisibleSheets(false, System.IO.Path.GetFileNameWithoutExtension(FileName));
                            PdfExport.EndExport();
                        }
                    }
                    catch
                    {
                        try
                        {
                            File.Delete(FileName);
                        }
                        catch
                        {
                            //Not here.
                        }
                        throw;
                    }
                }
                else
                {
                    PdfExport.PageLayout = TPageLayout.None;
                    PdfExport.Export(FileName);
                }
            }
            catch (Exception ex)
            {
                FMainException = ex;
            }
        }

        internal Exception MainException
        {
            get
            {
                return FMainException;
            }
        }
    }
    #endregion


}
