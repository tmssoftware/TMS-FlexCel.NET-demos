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



namespace ExportSVG
{
    /// <summary>
    /// An Example on how to export to SVG.
    /// </summary>
    public partial class mainForm: System.Windows.Forms.Form
    {
        private FlexCelSVGExport SVG = new FlexCelSVGExport();
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

        private void btnClose_Click(object sender, System.EventArgs e)
        {
            Close();
        }

        private void LoadSheetConfig()
        {
            ExcelFile Xls = SVG.Workbook;

            chGridLines.Checked = Xls.PrintGridLines;
            chPrintHeadings.Checked = Xls.PrintHeadings;
            chFormulaText.Checked = Xls.ShowFormulaText;
        }

        private void openFile_Click(object sender, System.EventArgs e)
        {
            if (openFileDialog1.ShowDialog() != DialogResult.OK) return;
            SVG.Workbook = new XlsFile();

            SVG.Workbook.Open(openFileDialog1.FileName);

            Text = "Export: " + openFileDialog1.FileName;

            ExcelFile Xls = SVG.Workbook;

            cbSheet.Items.Clear();
            for (int i = 1; i <= Xls.SheetCount; i++)
            {
                cbSheet.Items.Add(Xls.GetSheetName(i));
            }
            cbSheet.SelectedIndex = Xls.ActiveSheet - 1;

            LoadSheetConfig();
        }

        private bool CheckFileOpen()
        {
            if (SVG.Workbook == null)
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
                ExcelFile Xls = SVG.Workbook;

                //Note: In this demo we will only apply this things to the active sheet.
                //If you want to apply the settings to all the sheets, you should loop in the sheets and change them here.
                Xls.PrintGridLines = chGridLines.Checked;
                Xls.PrintHeadings = chPrintHeadings.Checked;
                Xls.ShowFormulaText = chFormulaText.Checked;

                SVG.PrintRangeLeft = Convert.ToInt32(edLeft.Text);
                SVG.PrintRangeTop = Convert.ToInt32(edTop.Text);
                SVG.PrintRangeRight = Convert.ToInt32(edRight.Text);
                SVG.PrintRangeBottom = Convert.ToInt32(edBottom.Text);

                SVG.HidePrintObjects = THidePrintObjects.None;
                if (!cbImages.Checked) SVG.HidePrintObjects |= THidePrintObjects.Images;
                if (!cbHyperlinks.Checked) SVG.HidePrintObjects |= THidePrintObjects.Hyperlynks;
                if (!cbComments.Checked) SVG.HidePrintObjects |= THidePrintObjects.Comments;
                if (!cbHeadersFooters.Checked) SVG.HidePrintObjects |= THidePrintObjects.HeadersAndFooters;

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
            SVG.Workbook.ActiveSheet = cbSheet.SelectedIndex + 1;
            LoadSheetConfig();
        }

        private void export_Click(object sender, System.EventArgs e)
        {
            if (!CheckFileOpen()) return;
            if (!LoadPreferences()) return;

            if (exportDialog.ShowDialog() != DialogResult.OK) return;

            SVG.AllowOverwritingFiles = true;

            SVG.AllVisibleSheets = cbExportObject.SelectedIndex == 0;

            SVG.SaveAsImage(
                   (x) =>
                    {
                        x.FileName = Path.ChangeExtension(exportDialog.FileName, "") + "_" + x.Workbook.SheetName + "_" + x.SheetPageNumber.ToString() + ".svg";
                    });

            if (MessageBox.Show("Do you want to open the folder with the generated files?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                Process.Start(Path.GetDirectoryName(exportDialog.FileName));
            }

        }


        private void mainForm_Load(object sender, System.EventArgs e)
        {
            cbExportObject.SelectedIndex = 0;
        }

        private void cbExportObject_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            cbSheet.Enabled = cbExportObject.SelectedIndex == 1;
        }
    }
}
