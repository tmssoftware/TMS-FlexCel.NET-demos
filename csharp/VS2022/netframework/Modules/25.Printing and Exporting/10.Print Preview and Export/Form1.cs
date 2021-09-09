using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using FlexCel.Core;
using FlexCel.XlsAdapter;
using FlexCel.Render;

using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;

using System.Runtime.InteropServices;
using FlexCel.Draw;

namespace PrintPreviewandExport
{
    /// <summary>
    /// Printing / Previewing and Exporting xls files.
    /// </summary>
    public partial class mainForm: System.Windows.Forms.Form
    {
        private FlexCel.Render.FlexCelPrintDocument flexCelPrintDocument1;

        public mainForm()
        {
            InitializeComponent();
            cbInterpolation.SelectedIndex = 1;
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
            ExcelFile Xls = flexCelPrintDocument1.Workbook;

            chGridLines.Checked = Xls.PrintGridLines;
            chHeadings.Checked = Xls.PrintHeadings;
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

            Landscape.Checked = (Xls.PrintOptions & TPrintOptions.Orientation) == 0;

        }

        private void openFile_Click(object sender, System.EventArgs e)
        {
            if (openFileDialog1.ShowDialog() != DialogResult.OK) return;
            flexCelPrintDocument1.Workbook = new XlsFile();

            flexCelPrintDocument1.Workbook.Open(openFileDialog1.FileName);

            edFileName.Text = openFileDialog1.FileName;

            ExcelFile Xls = flexCelPrintDocument1.Workbook;

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
            if (flexCelPrintDocument1.Workbook == null)
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
                flexCelPrintDocument1.AllVisibleSheets = cbAllSheets.Checked;
                flexCelPrintDocument1.ResetPageNumberOnEachSheet = cbResetPageNumber.Checked;
                flexCelPrintDocument1.AntiAliasedText = chAntiAlias.Checked;

                ExcelFile Xls = flexCelPrintDocument1.Workbook;
                Xls.PrintGridLines = chGridLines.Checked;
                Xls.PrintHeadings = chHeadings.Checked;
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


                flexCelPrintDocument1.PrintRangeLeft = Convert.ToInt32(edLeft.Text);
                flexCelPrintDocument1.PrintRangeTop = Convert.ToInt32(edTop.Text);
                flexCelPrintDocument1.PrintRangeRight = Convert.ToInt32(edRight.Text);
                flexCelPrintDocument1.PrintRangeBottom = Convert.ToInt32(edBottom.Text);

                flexCelPrintDocument1.DocumentName = flexCelPrintDocument1.Workbook.ActiveFileName + " - Sheet " + flexCelPrintDocument1.Workbook.ActiveSheetByName;

                flexCelPrintDocument1.DefaultPageSettings.Landscape = Landscape.Checked;
            }
            catch (Exception e)
            {
                MessageBox.Show("Error: " + e.Message);
                return false;
            }
            return true;
        }

        private void preview_Click(object sender, System.EventArgs e)
        {
            if (!HasFileOpen()) return;
            if (!LoadPreferences()) return;
            if (!DoSetup()) return;

            //If you want to bypass the paper size selected on the dialog and use the one on Excel, uncomment
            //the following lines:
            //TPaperDimensions t= flexCelPrintDocument1.Workbook.PrintPaperDimensions;
            //flexCelPrintDocument1.DefaultPageSettings.PaperSize = new PaperSize(t.PaperName, Convert.ToInt32(t.Width), Convert.ToInt32(t.Height));

            printPreviewDialog1.ShowDialog();

        }

        private bool DoSetup()
        {
            bool Result = printDialog1.ShowDialog() == DialogResult.OK;
            Landscape.Checked = flexCelPrintDocument1.DefaultPageSettings.Landscape;
            return Result;
        }

        private void setup_Click(object sender, System.EventArgs e)
        {
            DoSetup();
        }

        private void chFitIn_CheckedChanged(object sender, System.EventArgs e)
        {
            edVPages.ReadOnly = !chFitIn.Checked;
            edHPages.ReadOnly = !chFitIn.Checked;
            edZoom.ReadOnly = chFitIn.Checked;
        }

        private void print_Click(object sender, System.EventArgs e)
        {
            if (!HasFileOpen()) return;
            if (!LoadPreferences()) return;
            if (!DoSetup()) return;
            flexCelPrintDocument1.Print();
        }

        private void cbSheet_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            flexCelPrintDocument1.Workbook.ActiveSheet = cbSheet.SelectedIndex + 1;
            LoadSheetConfig();

        }


        /// <summary>
        /// Add a "Confidential" watermark on each page.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void flexCelPrintDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            if (!cbConfidential.Checked) return;

            using (Matrix myMatrix = new Matrix())
            {
                myMatrix.RotateAt(45, new PointF(e.PageBounds.Left + e.MarginBounds.Width / 2F, e.PageBounds.Top + e.MarginBounds.Height / 2F), MatrixOrder.Append);
                e.Graphics.Transform = myMatrix;
            }

            using (Brush ABrush = new SolidBrush(Color.FromArgb(30, 25, 25, 25))) //Red=Green=Blue is a shade of gray. Alpha=30 means it is transparent (255 is pure opaque, 0 is pure transparent).
            {
                using (Font AFont = new Font("Arial", 72))
                {
                    using (StringFormat sf = new StringFormat())
                    {

                        sf.Alignment = StringAlignment.Center;
                        sf.LineAlignment = StringAlignment.Center;
                        e.Graphics.DrawString("Confidential", AFont, ABrush, e.PageBounds, sf);
                    }
                }
            }
        }

        #region Hard Margins
        //Shows how to read the hard margins from a printer if you really need to.

        private void flexCelPrintDocument1_BeforePrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            switch (cbInterpolation.SelectedIndex)
            {
                case 0: e.Graphics.InterpolationMode = InterpolationMode.Bicubic; break;
                case 1: e.Graphics.InterpolationMode = InterpolationMode.Bilinear; break;
                case 2: e.Graphics.InterpolationMode = InterpolationMode.Default; break;
                case 3: e.Graphics.InterpolationMode = InterpolationMode.High; break;
                case 4: e.Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic; break;
                case 5: e.Graphics.InterpolationMode = InterpolationMode.HighQualityBilinear; break;
                case 6: e.Graphics.InterpolationMode = InterpolationMode.Low; break;
                case 7: e.Graphics.InterpolationMode = InterpolationMode.NearestNeighbor; break;
            }

        }

        [DllImport("gdi32.dll")]
        private static extern Int32 GetDeviceCaps(IntPtr hdc, Int32 capindex);

        /// <summary>
        /// This event will adjust for a better position on the page for some printers. 
        /// It is not normally necessary, and it has to make an unmanaged call to GetDeviceCaps,
        /// but it is given here as an example of how it could be done.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void flexCelPrintDocument1_GetPrinterHardMargins(object sender, FlexCel.Render.PrintHardMarginsEventArgs e)
        {
            const int PHYSICALOFFSETX = 112;
            const int PHYSICALOFFSETY = 113;

            double DpiX = e.Graphics.DpiX;
            double DpiY = e.Graphics.DpiY;

            IntPtr Hdc = e.Graphics.GetHdc();
            try
            {
                e.XMargin = (float)(GetDeviceCaps(Hdc, PHYSICALOFFSETX) * 100.0 / DpiX);
                e.YMargin = (float)(GetDeviceCaps(Hdc, PHYSICALOFFSETY) * 100.0 / DpiY);
            }

            finally
            {
                e.Graphics.ReleaseHdc(Hdc);
            }

        }
        #endregion

        #region Export as image

        #region Common methods to Export with FlexCelImgExport
        private Bitmap CreateBitmap(double Resolution, TPaperDimensions pd, PixelFormat PxFormat)
        {
            Bitmap Result =
                new Bitmap((int)Math.Ceiling(pd.Width / 96F * Resolution),
                (int)Math.Ceiling(pd.Height / 96F * Resolution), PxFormat);
            Result.SetResolution((float)Resolution, (float)Resolution);
            return Result;

        }
        #endregion

        #region Export using FlexCelImgExport - simple images the hard way. DO NOT USE IF NOT DESPERATE!
        //The methods shows how to use FlexCelImgExport the "hard way", without using SaveAsImage.
        //For normal operation you should only need to call SaveAsImage, but you could use the code here
        //if you need to customize the ImgExport output, or if you need to get all the images as different files.
        private void CreateImg(Stream OutStream, FlexCelImgExport ImgExport, ImageFormat ImgFormat, ImageColorDepth Colors, ref TImgExportInfo ExportInfo)
        {
            TPaperDimensions pd = ImgExport.GetRealPageSize();

            PixelFormat RgbPixFormat;
            if (Colors != ImageColorDepth.TrueColor) RgbPixFormat = PixelFormat.Format32bppPArgb; else RgbPixFormat = PixelFormat.Format24bppRgb;
            PixelFormat PixFormat = PixelFormat.Format1bppIndexed;
            switch (Colors)
            {
                case ImageColorDepth.TrueColor: PixFormat = RgbPixFormat; break;
                case ImageColorDepth.Color256: PixFormat = PixelFormat.Format8bppIndexed; break;
            }

            using (Bitmap OutImg = CreateBitmap(ImgExport.Resolution, pd, PixFormat))
            {
                Bitmap ActualOutImg;
                if (Colors != ImageColorDepth.TrueColor) ActualOutImg = CreateBitmap(ImgExport.Resolution, pd, RgbPixFormat); else ActualOutImg = OutImg;
                try
                {
                    using (Graphics Gr = Graphics.FromImage(ActualOutImg))
                    {
                        Gr.FillRectangle(Brushes.White, 0, 0, ActualOutImg.Width, ActualOutImg.Height); //Clear the background
                        ImgExport.ExportNext(Gr, ref ExportInfo);
                    }

                    if (Colors == ImageColorDepth.BlackAndWhite) FloydSteinbergDither.ConvertToBlackAndWhite(ActualOutImg, OutImg);
                    else
                        if (Colors == ImageColorDepth.Color256)
                    {
                        OctreeQuantizer.ConvertTo256Colors(ActualOutImg, OutImg);
                    }
                }
                finally
                {
                    if (ActualOutImg != OutImg) ActualOutImg.Dispose();
                }

                OutImg.Save(OutStream, ImgFormat);
            }
        }

        private void ExportAllImages(FlexCelImgExport ImgExport, ImageFormat ImgFormat, ImageColorDepth ColorDepth)
        {
            TImgExportInfo ExportInfo = null; //For first page.
            int i = 0;
            do
            {
                string FileName = Path.GetDirectoryName(exportImageDialog.FileName)
                    + Path.DirectorySeparatorChar
                    + Path.GetFileNameWithoutExtension(exportImageDialog.FileName)
                    + "_" + ImgExport.Workbook.SheetName
                    + String.Format("_{0:0000}", i) +
                    Path.GetExtension(exportImageDialog.FileName);
                using (FileStream ImageStream = new FileStream(FileName, FileMode.Create))
                {
                    CreateImg(ImageStream, ImgExport, ImgFormat, ColorDepth, ref ExportInfo);
                }
                i++;
            } while (ExportInfo.CurrentPage < ExportInfo.TotalPages);
        }

        private void DoExportUsingFlexCelImgExportComplex(ImageColorDepth ColorDepth)
        {
            if (!HasFileOpen()) return;
            if (!LoadPreferences()) return;

            if (exportImageDialog.ShowDialog() != DialogResult.OK) return;

            System.Drawing.Imaging.ImageFormat ImgFormat = System.Drawing.Imaging.ImageFormat.Png;
            if (String.Compare(Path.GetExtension(exportImageDialog.FileName), ".jpg", true) == 0)
                ImgFormat = System.Drawing.Imaging.ImageFormat.Jpeg;

            using (FlexCelImgExport ImgExport = new FlexCelImgExport(flexCelPrintDocument1.Workbook))
            {
                ImgExport.Resolution = 96; //To get a better quality image but with larger file size too, increate this value. (for example to 300 or 600 dpi)

                if (cbAllSheets.Checked)
                {
                    int SaveActiveSheet = ImgExport.Workbook.ActiveSheet;
                    try
                    {
                        ImgExport.Workbook.ActiveSheet = 1;
                        bool Finished = false;
                        while (!Finished)
                        {
                            ExportAllImages(ImgExport, ImgFormat, ColorDepth);
                            if (ImgExport.Workbook.ActiveSheet < ImgExport.Workbook.SheetCount)
                            {
                                ImgExport.Workbook.ActiveSheet++;
                            }
                            else Finished = true;

                        }
                    }
                    finally
                    {
                        ImgExport.Workbook.ActiveSheet = SaveActiveSheet;
                    }
                }
                else
                {
                    ExportAllImages(ImgExport, ImgFormat, ColorDepth);
                }

            }

        }
        #endregion

        #region Export using FlexCelImgExport - simple images the simple way.

        private void DoExportUsingFlexCelImgExportSimple(ImageColorDepth ColorDepth)
        {
            if (!HasFileOpen()) return;
            if (!LoadPreferences()) return;

            if (exportImageDialog.ShowDialog() != DialogResult.OK) return;

            ImageExportType ImgFormat = ImageExportType.Png;
            if (String.Compare(Path.GetExtension(exportImageDialog.FileName), ".jpg", true) == 0)
                ImgFormat = ImageExportType.Jpeg;

            using (FlexCelImgExport ImgExport = new FlexCelImgExport(flexCelPrintDocument1.Workbook))
            {
                ImgExport.AllVisibleSheets = cbAllSheets.Checked;
                ImgExport.ResetPageNumberOnEachSheet = cbResetPageNumber.Checked;
                ImgExport.Resolution = 96; //To get a better quality image but with larger file size too, increate this value. (for example to 300 or 600 dpi)
                ImgExport.SaveAsImage(exportImageDialog.FileName, ImgFormat, ColorDepth);
            }
        }

        #endregion

        #region Export using FlexCelImageExport - MultiPageTiff
        //How to create a multipage tiff using FlexCelImgExport.        
        //This will create a multipage tiff with the data.
        private void DoExportMultiPageTiff(ImageColorDepth ColorDepth, bool IsFax)
        {
            if (!HasFileOpen()) return;
            if (!LoadPreferences()) return;

            if (exportTiffDialog.ShowDialog() != DialogResult.OK) return;

            ImageExportType ExportType = ImageExportType.Tiff;
            if (IsFax) ExportType = ImageExportType.Fax;

            using (FlexCelImgExport ImgExport = new FlexCelImgExport(flexCelPrintDocument1.Workbook))
            {
                ImgExport.AllVisibleSheets = cbAllSheets.Checked;
                ImgExport.ResetPageNumberOnEachSheet = cbResetPageNumber.Checked;

                ImgExport.Resolution = 96; //To get a better quality image but with larger file size too, increate this value. (for example to 300 or 600 dpi)
                using (FileStream TiffStream = new FileStream(exportTiffDialog.FileName, FileMode.Create))
                {
                    ImgExport.SaveAsImage(TiffStream, ExportType, ColorDepth);
                }
            }
            if (MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                Process.Start(exportTiffDialog.FileName);
            }

        }
        #endregion

        #region Event handlers
        private void ImgBlackAndWhite_Click(object sender, System.EventArgs e)
        {
            DoExportUsingFlexCelImgExportComplex(ImageColorDepth.BlackAndWhite);
        }

        private void Img256Colors_Click(object sender, System.EventArgs e)
        {
            DoExportUsingFlexCelImgExportComplex(ImageColorDepth.Color256);
        }

        private void ImgTrueColor_Click(object sender, System.EventArgs e)
        {
            DoExportUsingFlexCelImgExportComplex(ImageColorDepth.TrueColor);
        }

        private void ImgBlackAndWhite2_Click(object sender, System.EventArgs e)
        {
            DoExportUsingFlexCelImgExportSimple(ImageColorDepth.BlackAndWhite);
        }

        private void Img256Colors2_Click(object sender, System.EventArgs e)
        {
            DoExportUsingFlexCelImgExportSimple(ImageColorDepth.Color256);
        }

        private void ImgTrueColor2_Click(object sender, System.EventArgs e)
        {
            DoExportUsingFlexCelImgExportSimple(ImageColorDepth.TrueColor);
        }

        private void TiffFax_Click(object sender, System.EventArgs e)
        {
            DoExportMultiPageTiff(ImageColorDepth.BlackAndWhite, true);
        }

        private void TiffBlackAndWhite_Click(object sender, System.EventArgs e)
        {
            DoExportMultiPageTiff(ImageColorDepth.BlackAndWhite, false);
        }

        private void Tiff256Colors_Click(object sender, System.EventArgs e)
        {
            DoExportMultiPageTiff(ImageColorDepth.Color256, false);
        }

        private void TiffTrueColor_Click(object sender, System.EventArgs e)
        {
            DoExportMultiPageTiff(ImageColorDepth.TrueColor, false);
        }

        #endregion

        private void cbAllSheets_CheckedChanged(object sender, System.EventArgs e)
        {
            cbSheet.Enabled = !cbAllSheets.Checked;
            cbResetPageNumber.Enabled = cbAllSheets.Checked;
            Landscape.Enabled = !cbAllSheets.Checked;  //When exporting many sheets, we will honor the landscape/portrait setting on each one.
        }



        #endregion

    }
}
