using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using FlexCel.Render;
using System.IO;
using System.Diagnostics;

using System.Text;



namespace ExportHTML
{
    /// <summary>
    /// An Example on how to export to HTML.
    /// </summary>
    public partial class mainForm: System.Windows.Forms.Form
    {
        private FlexCel.Render.FlexCelHtmlExport flexCelHtmlExport1;


        public mainForm()
        {
            InitializeComponent();
        }

        private Mailform MailDialog;

        private void button2_Click(object sender, System.EventArgs e)
        {
            Close();
        }

        private void LoadSheetConfig()
        {
            ExcelFile Xls = flexCelHtmlExport1.Workbook;

            chGridLines.Checked = Xls.PrintGridLines;
            chPrintHeadings.Checked = Xls.PrintHeadings;
            chFormulaText.Checked = Xls.ShowFormulaText;
        }

        private void openFile_Click(object sender, System.EventArgs e)
        {
            if (openFileDialog1.ShowDialog() != DialogResult.OK) return;
            flexCelHtmlExport1.Workbook = new XlsFile();

            flexCelHtmlExport1.Workbook.Open(openFileDialog1.FileName);

            Text = "Export: " + openFileDialog1.FileName;

            ExcelFile Xls = flexCelHtmlExport1.Workbook;

            cbSheet.Items.Clear();
            for (int i = 1; i <= Xls.SheetCount; i++)
            {
                cbSheet.Items.Add(Xls.GetSheetName(i));
            }
            cbSheet.SelectedIndex = Xls.ActiveSheet - 1;

            LoadSheetConfig();
        }

        private bool HasFileOpen()
        {
            if (flexCelHtmlExport1.Workbook == null)
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
                ExcelFile Xls = flexCelHtmlExport1.Workbook;

                //Note: In this demo we will only apply this things to the active sheet.
                //If you want to apply the settings to all the sheets, you should loop in the sheets and change them here.
                Xls.PrintGridLines = chGridLines.Checked;
                Xls.PrintHeadings = chPrintHeadings.Checked;
                Xls.ShowFormulaText = chFormulaText.Checked;

                flexCelHtmlExport1.PrintRangeLeft = Convert.ToInt32(edLeft.Text);
                flexCelHtmlExport1.PrintRangeTop = Convert.ToInt32(edTop.Text);
                flexCelHtmlExport1.PrintRangeRight = Convert.ToInt32(edRight.Text);
                flexCelHtmlExport1.PrintRangeBottom = Convert.ToInt32(edBottom.Text);

                if (sbSVG.Checked) flexCelHtmlExport1.SavedImagesFormat = THtmlImageFormat.Svg; else flexCelHtmlExport1.SavedImagesFormat = THtmlImageFormat.Png;
                flexCelHtmlExport1.EmbedImages = cbEmbedImages.Checked;

                flexCelHtmlExport1.FixOutlook2007CssSupport = cbOutlook2007.Checked;
                flexCelHtmlExport1.FixIE6TransparentPngSupport = cbIe6Png.Checked;

                flexCelHtmlExport1.HidePrintObjects = THidePrintObjects.None;
                if (!cbImages.Checked) flexCelHtmlExport1.HidePrintObjects |= THidePrintObjects.Images;
                if (!cbHyperlinks.Checked) flexCelHtmlExport1.HidePrintObjects |= THidePrintObjects.Hyperlynks;
                if (!cbComments.Checked) flexCelHtmlExport1.HidePrintObjects |= THidePrintObjects.Comments;
                if (!cbHeadersFooters.Checked) flexCelHtmlExport1.HidePrintObjects |= THidePrintObjects.HeadersAndFooters;

            }
            catch (Exception e)
            {
                MessageBox.Show("Error: " + e.Message);
                return false;
            }
            return true;
        }


        private void cbSheet_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            flexCelHtmlExport1.Workbook.ActiveSheet = cbSheet.SelectedIndex + 1;
            LoadSheetConfig();
        }

        private void export_Click(object sender, System.EventArgs e)
        {
            if (!HasFileOpen()) return;
            if (!LoadPreferences()) return;

            if (cbFileFormat.SelectedIndex == 1)
            {
                flexCelHtmlExport1.HtmlFileFormat = THtmlFileFormat.MHtml;
                exportDialog.FilterIndex = 2;
            }
            else
            {
                flexCelHtmlExport1.HtmlFileFormat = THtmlFileFormat.Html;
                exportDialog.FilterIndex = 1;
            }

            if (exportDialog.ShowDialog() != DialogResult.OK) return;

            flexCelHtmlExport1.AllowOverwritingFiles = true;

            string CssFileName = null;
            if (cbCss.Checked) CssFileName = edCss.Text;

            string FileNameToOpen = exportDialog.FileName;

            switch (cbHtmlVersion.SelectedIndex)
            {
                case 0: flexCelHtmlExport1.HtmlVersion = THtmlVersion.Html_32; break;
                case 2: flexCelHtmlExport1.HtmlVersion = THtmlVersion.XHTML_10; break;
                case 3: flexCelHtmlExport1.HtmlVersion = THtmlVersion.Html_5; break;
                default:
                    flexCelHtmlExport1.HtmlVersion = THtmlVersion.Html_401;
                    break;
            }

            if (edBodyStart.Text != null) flexCelHtmlExport1.ExtraInfo.BodyStart = new string[] { edBodyStart.Text };

            switch (cbExportObject.SelectedIndex)
            {
                case 0:
                    TSheetSelectorPosition SelectorPosition = TSheetSelectorPosition.None;

                    //If in VB.NET or Delphi.NET, use "if cbTop.Checked then SelectorPosition = SelectorPosition or TSheetSelectorPosition.Top"
                    if (cbTop.Checked) SelectorPosition |= TSheetSelectorPosition.Top;
                    if (cbLeft.Checked) SelectorPosition |= TSheetSelectorPosition.Left;
                    if (cbBottom.Checked) SelectorPosition |= TSheetSelectorPosition.Bottom;
                    if (cbRight.Checked) SelectorPosition |= TSheetSelectorPosition.Right;


                    flexCelHtmlExport1.ExportAllVisibleSheetsAsTabs(Path.GetDirectoryName(exportDialog.FileName),
                        Path.GetFileNameWithoutExtension(exportDialog.FileName), Path.GetExtension(exportDialog.FileName), edImages.Text, CssFileName, new TStandardSheetSelector(SelectorPosition));

                    FileNameToOpen = Path.Combine(Path.GetDirectoryName(exportDialog.FileName), Path.GetFileNameWithoutExtension(exportDialog.FileName));
                    FileNameToOpen = Path.Combine(FileNameToOpen, flexCelHtmlExport1.Workbook.SheetName);
                    FileNameToOpen = Path.Combine(FileNameToOpen, Path.GetExtension(exportDialog.FileName));

                    break;
                case 1:
                    flexCelHtmlExport1.ExportAllVisibleSheetsAsOneHtmlFile(exportDialog.FileName, edImages.Text, CssFileName, edSheetSeparator.Text);
                    break;

                case 2:
                    {
                        flexCelHtmlExport1.Export(exportDialog.FileName, edImages.Text, CssFileName);
                        break;
                    }
            }

            string[] GeneratedFiles = flexCelHtmlExport1.GeneratedFiles.GetHtmlFiles();
            if (GeneratedFiles.Length == 0)
            {
                MessageBox.Show("Error: No file has been generated");
            }
            else
            {
                if (MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    Process.Start(GeneratedFiles[0]);
                }
            }
        }

        private void btnEmail_Click(object sender, System.EventArgs e)
        {
            if (!HasFileOpen()) return;
            if (MailDialog == null) MailDialog = new Mailform();
            MailDialog.MainForm = this;

            if (!flexCelHtmlExport1.FixOutlook2007CssSupport)
            {
                DialogResult dr = MessageBox.Show("You have not checked \"Outlook 2007 support\". If any of your clients has Outlook express, you should turn this on.\n\nUse Outlook 2007 fix?", "Warning", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);

                if (dr == DialogResult.Cancel) return;
                if (dr == DialogResult.Yes)
                {
                    cbOutlook2007.Checked = true;
                    flexCelHtmlExport1.FixOutlook2007CssSupport = true;
                }
            }

            MailDialog.ShowDialog();

        }

        public byte[] GenerateMHTML()
        {
            LoadPreferences();
            flexCelHtmlExport1.HtmlFileFormat = THtmlFileFormat.MHtml;


            flexCelHtmlExport1.AllowOverwritingFiles = true;

            flexCelHtmlExport1.HtmlVersion = THtmlVersion.Html_401;

            if (edBodyStart.Text != null) flexCelHtmlExport1.ExtraInfo.BodyStart = new string[] { edBodyStart.Text };

            using (MemoryStream ms = new MemoryStream())
            {
                using (StreamWriter writer = new StreamWriter(ms, Encoding.UTF8))
                {
                    flexCelHtmlExport1.Export(writer, flexCelHtmlExport1.Workbook.ActiveFileName, null);
                }
                return ms.ToArray();
            }
        }


        private void mainForm_Load(object sender, System.EventArgs e)
        {
            cbExportObject.SelectedIndex = 0;
            cbHtmlVersion.SelectedIndex = 3;
            cbFileFormat.SelectedIndex = 0;
        }

        private void cbExportObject_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            edSheetSeparator.Enabled = cbExportObject.SelectedIndex == 1;
            cbTop.Enabled = cbExportObject.SelectedIndex == 0;
            cbLeft.Enabled = cbExportObject.SelectedIndex == 0;
            cbRight.Enabled = cbExportObject.SelectedIndex == 0;
            cbBottom.Enabled = cbExportObject.SelectedIndex == 0;
            cbSheet.Enabled = cbExportObject.SelectedIndex == 2;
        }

        private void cbCss_CheckedChanged(object sender, System.EventArgs e)
        {
            edCss.Enabled = cbCss.Checked;
        }

        private void flexCelHtmlExport1_HtmlFont(object sender, FlexCel.Core.HtmlFontEventArgs e)
        {
            if (cbReplaceFonts.Checked)
            {
                e.FontFamily = "arial, sans-serif";
            }
        }

    }
}
