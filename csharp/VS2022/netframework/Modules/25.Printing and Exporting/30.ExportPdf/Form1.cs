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
using System.Drawing.Drawing2D;
using FlexCel.Pdf;

//only needed if you want to go unmanaged.
using System.Runtime.InteropServices;


namespace ExportPdf
{
    /// <summary>
    /// Exporting xls files to pdf.
    /// </summary>
    public partial class mainForm: System.Windows.Forms.Form
    {
        private FlexCel.Render.FlexCelPdfExport flexCelPdfExport1;

        public mainForm()
        {
            InitializeComponent();
            cbFontMapping.SelectedIndex = 1;
            cbPdfType.SelectedIndex = 0;
            cbTagged.SelectedIndex = 0;
            cbVersion.SelectedIndex = 1;
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

        private void LoadSheetConfig()
        {
            ExcelFile Xls = flexCelPdfExport1.Workbook;

            chGridLines.Checked = Xls.PrintGridLines;
            chFormulaText.Checked = Xls.ShowFormulaText;

            chPrintLeft.Checked = (Xls.PrintOptions & TPrintOptions.LeftToRight) != 0;
            edHeader.Text = Xls.PageHeader;
            edFooter.Text = Xls.PageFooter;
            chFitIn.Checked = Xls.PrintToFit;
            edHPages.Text = Xls.PrintNumberOfHorizontalPages.ToString();
            edVPages.Text = Xls.PrintNumberOfVerticalPages.ToString();
            edVPages.ReadOnly = !chFitIn.Checked;
            edHPages.ReadOnly = !chFitIn.Checked;

            edZoom.ReadOnly = chFitIn.Checked;
            edZoom.Text = Xls.PrintScale.ToString();

            TXlsMargins m = Xls.GetPrintMargins();
            edl.Text = m.Left.ToString();
            edt.Text = m.Top.ToString();
            edr.Text = m.Right.ToString();
            edb.Text = m.Bottom.ToString();
            edf.Text = m.Footer.ToString();
            edh.Text = m.Header.ToString();

            chLandscape.Checked = Xls.PrintLandscape;

            edAuthor.Text = Convert.ToString(Xls.DocumentProperties.GetStandardProperty(TPropertyId.Author));
            edTitle.Text = Convert.ToString(Xls.DocumentProperties.GetStandardProperty(TPropertyId.Title));
            edSubject.Text = Convert.ToString(Xls.DocumentProperties.GetStandardProperty(TPropertyId.Subject));
        }

        private void openFile_Click(object sender, System.EventArgs e)
        {
            if (openFileDialog1.ShowDialog() != DialogResult.OK) return;
            flexCelPdfExport1.Workbook = new XlsFile();

            flexCelPdfExport1.Workbook.Open(openFileDialog1.FileName);

            edFileName.Text = openFileDialog1.FileName;

            ExcelFile Xls = flexCelPdfExport1.Workbook;

            cbSheet.Items.Clear();
            int ActSheet = Xls.ActiveSheet;
            for (int i = 1; i <= Xls.SheetCount; i++)
            {
                Xls.ActiveSheet = i;
                cbSheet.Items.Add(Xls.SheetName);
            }
            Xls.ActiveSheet = ActSheet;
            cbSheet.SelectedIndex = ActSheet - 1;

            LoadSheetConfig();
        }

        private bool HasFileOpen()
        {
            if (flexCelPdfExport1.Workbook == null)
            {
                MessageBox.Show("You need to open a file first.");
                return false;
            }
            return true;
        }

        private bool LoadPreferences()
        {
            //NOTE: THERE SHOULD BE *A LOT* MORE VALIDATION OF VALUES ON THIS METHOD. (For example, validate that margins are between bounds)
            // As this is a simple demo, they are not included. 
            try
            {
                ExcelFile Xls = flexCelPdfExport1.Workbook;
                Xls.PrintGridLines = chGridLines.Checked;
                Xls.PageHeader = edHeader.Text;
                Xls.PageFooter = edFooter.Text;
                Xls.ShowFormulaText = chFormulaText.Checked;

                if (chFitIn.Checked)
                {
                    Xls.PrintToFit = true;
                    Xls.PrintNumberOfHorizontalPages = Convert.ToInt32(edHPages.Text);
                    Xls.PrintNumberOfVerticalPages = Convert.ToInt32(edVPages.Text);
                }
                else
                    Xls.PrintToFit = false;

                if (chPrintLeft.Checked) Xls.PrintOptions |= TPrintOptions.LeftToRight;
                else Xls.PrintOptions &= ~TPrintOptions.LeftToRight;

                try
                {
                    Xls.PrintScale = Convert.ToInt32(edZoom.Text);
                }
                catch
                {
                    MessageBox.Show("Invalid Zoom");
                    return false;
                }

                TXlsMargins m = new TXlsMargins();
                m.Left = Convert.ToDouble(edl.Text);
                m.Top = Convert.ToDouble(edt.Text);
                m.Right = Convert.ToDouble(edr.Text);
                m.Bottom = Convert.ToDouble(edb.Text);
                m.Footer = Convert.ToDouble(edf.Text);
                m.Header = Convert.ToDouble(edh.Text);
                Xls.SetPrintMargins(m);

                flexCelPdfExport1.PrintRangeLeft = Convert.ToInt32(edLeft.Text);
                flexCelPdfExport1.PrintRangeTop = Convert.ToInt32(edTop.Text);
                flexCelPdfExport1.PrintRangeRight = Convert.ToInt32(edRight.Text);
                flexCelPdfExport1.PrintRangeBottom = Convert.ToInt32(edBottom.Text);

                if (chEmbed.Checked)
                    flexCelPdfExport1.FontEmbed = TFontEmbed.Embed;
                else flexCelPdfExport1.FontEmbed = TFontEmbed.None;

                if (chSubset.Checked)
                    flexCelPdfExport1.FontSubset = TFontSubset.Subset;
                else flexCelPdfExport1.FontSubset = TFontSubset.DontSubset;

                flexCelPdfExport1.Kerning = cbKerning.Checked;

                switch (cbFontMapping.SelectedIndex)
                {
                    case 0: flexCelPdfExport1.FontMapping = TFontMapping.ReplaceAllFonts; break;
                    case 1: flexCelPdfExport1.FontMapping = TFontMapping.ReplaceStandardFonts; break;
                    case 2: flexCelPdfExport1.FontMapping = TFontMapping.DontReplaceFonts; break;
                }

                switch (cbPdfType.SelectedIndex)
                {
                    case 0:
                        flexCelPdfExport1.PdfType = TPdfType.Standard;
                        break;
                    case 1:
                        flexCelPdfExport1.PdfType = TPdfType.PDFA1;
                        break;
                    case 2:
                        flexCelPdfExport1.PdfType = TPdfType.PDFA2;
                        break;
                    case 3:
                        flexCelPdfExport1.PdfType = TPdfType.PDFA3;
                        break;
                }

                switch (cbTagged.SelectedIndex)
                {
                    case 0: flexCelPdfExport1.TagMode = TTagMode.Full; break;
                    case 1: flexCelPdfExport1.TagMode = TTagMode.None; break;
                }

                switch (cbVersion.SelectedIndex)
                {
                    case 0: flexCelPdfExport1.PdfVersion = TPdfVersion.v14; break;
                    case 1: flexCelPdfExport1.PdfVersion = TPdfVersion.v16; break;
                }

                flexCelPdfExport1.Properties.Author = edAuthor.Text;
                flexCelPdfExport1.Properties.Title = edTitle.Text;
                flexCelPdfExport1.Properties.Subject = edSubject.Text;
                flexCelPdfExport1.Properties.Language = edLang.Text;

                Xls.PrintLandscape = chLandscape.Checked;
            }
            catch (Exception e)
            {
                MessageBox.Show("Error: " + e.Message);
                return false;
            }
            return true;
        }


        private void chFitIn_CheckedChanged(object sender, System.EventArgs e)
        {
            edVPages.ReadOnly = !chFitIn.Checked;
            edHPages.ReadOnly = !chFitIn.Checked;
            edZoom.ReadOnly = chFitIn.Checked;
        }

        private void cbSheet_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            flexCelPdfExport1.Workbook.ActiveSheet = cbSheet.SelectedIndex + 1;
            LoadSheetConfig();
        }

        private void export_Click(object sender, System.EventArgs e)
        {
            if (!HasFileOpen()) return;
            if (!LoadPreferences()) return;

            if (exportDialog.ShowDialog() != DialogResult.OK) return;

            using (FileStream Pdf = new FileStream(exportDialog.FileName, FileMode.Create))
            {
                int SaveSheet = flexCelPdfExport1.Workbook.ActiveSheet;
                try
                {
                    flexCelPdfExport1.BeginExport(Pdf);
                    if (chExportAll.Checked)
                    {
                        flexCelPdfExport1.PageLayout = TPageLayout.Outlines; //To how the bookmarks when opening the file.
                        flexCelPdfExport1.ExportAllVisibleSheets(cbResetPageNumber.Checked, Path.GetFileNameWithoutExtension(exportDialog.FileName));
                    }
                    else
                    {
                        flexCelPdfExport1.PageLayout = TPageLayout.None;
                        flexCelPdfExport1.ExportSheet();
                    }
                    flexCelPdfExport1.EndExport();
                }
                finally
                {
                    flexCelPdfExport1.Workbook.ActiveSheet = SaveSheet;
                }

                if (MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    Process.Start(exportDialog.FileName);
                }
            }
        }

        private void chExportAll_CheckedChanged(object sender, System.EventArgs e)
        {
            cbSheet.Enabled = !chExportAll.Checked;
            cbResetPageNumber.Enabled = chExportAll.Checked;
        }

        /// <summary>
        /// Add a "Confidential" watermark on each page.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void flexCelPdfExport1_AfterGeneratePage(object sender, FlexCel.Render.PageEventArgs e)
        {
            if (!cbConfidential.Checked) return;

            const string s = "Confidential";

            using (Brush ABrush = new SolidBrush(Color.FromArgb(30, 25, 25, 25))) //Red=Green=Blue is a shade of gray. Alpha=30 means it is transparent (255 is pure opaque, 0 is pure transparent).
            {
                using (TUIFont AFont = TUIFont.Create("Arial", 72))
                {
                    double x0 = e.File.PageSize.Width * 72.0 / 100.0 / 2.0;   //PageSize is in inches/100, our coordinate system is in Points, that is inches/72
                    double y0 = e.File.PageSize.Height * 72.0 / 100.0 / 2.0;
                    SizeF sf = e.File.MeasureString(s, AFont);
                    e.File.Rotate(x0, y0, 45);
                    e.File.DrawString(s, AFont, ABrush, x0 - sf.Width / 2.0, y0 + sf.Height / 2.0);  //the y coord means the bottom of the text, and as the y axis grows down, we have to add sf.height/2 instead of substracting it.
                }
            }
        }


        /// <summary>
        /// We show on this event how you can make an unmanaged call to the Win32 API to return font information and avoid
        /// scanning the "fonts" folder. Note that this is <b>UNMANAGED</b> code, and it is not really needed except for small performance concerns,
        /// so avoid using it if you don't really need it. Please read UsingFlexCelPdfExport for more information.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void flexCelPdfExport1_GetFontData(object sender, FlexCel.Pdf.GetFontDataEventArgs e)
        {
            //If the checkbox is not checked, just ignore this event.
            if (!cbUseGetFontData.Checked)
            {
                e.Applied = false;
                return;
            }

            //Actually make the WIN32 call.
            uint ttcf = 0x66637474; //return full true type collections.

            // Allocate a handle for the font
            IntPtr FontHandle = ((FlexCel.Draw.TGdipUIFont)e.InputFont).Handle.ToHfont();
            try
            {
                using (Graphics Gr = Graphics.FromHwnd(IntPtr.Zero))
                {
                    IntPtr GrHandle = Gr.GetHdc();
                    try
                    {
                        IntPtr ObjHandle = SelectObject(GrHandle, FontHandle);
                        try
                        {
                            //First find out the sizes
                            uint Size = GetFontData(GrHandle, ttcf, 0, null, 0);
                            if ((int)Size < 0) //error
                            {
                                ttcf = 0; //This might not be a true type collection, try again.
                                Size = GetFontData(GrHandle, ttcf, 0, null, 0);

                                if ((int)Size < 0) //nothing else to do, exit.
                                {
                                    e.Applied = false;
                                    return;
                                }
                            }

                            //Now get the font data.
                            e.FontData = new byte[(int)Size];
                            uint Result = GetFontData(GrHandle, ttcf, 0, e.FontData, Size);

                            if ((int)Result < 0)
                            {
                                e.Applied = false;
                                return;
                            }
                            e.Applied = true;
                        }
                        finally
                        {
                            DeleteObject(ObjHandle);
                        }
                    }
                    finally
                    {
                        Gr.ReleaseHdc(GrHandle);
                    }
                }
            }
            finally
            {
                DeleteObject(FontHandle);
            }
        }

        /// <summary>
        /// The Win32 call.
        /// </summary>
        /// <param name="hdc"></param>
        /// <param name="dwTable"></param>
        /// <param name="dwOffset"></param>
        /// <param name="lpvBuffer"></param>
        /// <param name="cbData"></param>
        /// <returns></returns>
        [DllImport("gdi32.dll")]
        static extern uint GetFontData(IntPtr hdc, uint dwTable, uint dwOffset,
            [In, Out] byte[] lpvBuffer, uint cbData);

        [DllImport("GDI32.dll")]
        static extern bool DeleteObject(IntPtr objectHandle);

        [DllImport("gdi32.dll")]
        static extern IntPtr SelectObject(IntPtr hdc, IntPtr hgdiobj);




    }
}
