using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using FlexCel.Core;
using FlexCel.XlsAdapter;
using System.Collections.Generic;

namespace VirtualMode
{
    /// <summary>
    /// A demo on how to read a file from FlexCel and display the results.
    /// </summary>
    public partial class mainForm: System.Windows.Forms.Form
    {
        SparseCellArray CellData; //we will store the data here. This is an example, in real world you would use "Virtual mode" to load the cells into your own structures.

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

        private void btnExit_Click(object sender, System.EventArgs e)
        {
            Close();
        }

        private void btnInfo_Click(object sender, System.EventArgs e)
        {
            MessageBox.Show("This demo shows how to read the contents of an xls file without loading the file in memory.\n" +
                "We will first load the sheet names in the file, then open just a single sheet, and read all or just the 50 first rows of it."
                );
        }


        private void btnOpenFile_Click(object sender, System.EventArgs e)
        {
            if (openFileDialog1.ShowDialog() != DialogResult.OK) return;
            ImportFile(openFileDialog1.FileName);
        }

        private void ImportFile(string FileName)
        {
            try
            {
                XlsFile xls = new XlsFile();
                xls.VirtualMode = true; //Remember to turn virtual mode on, or the event won't be called.

                //By default, FlexCel returns the formula text for the formulas, besides its calculated value.
                //If you are not interested in formula texts, you can gain a little performance by ignoring it.
                //This also works in non virtual mode.
                xls.IgnoreFormulaText = cbIgnoreFormulaText.Checked;

                CellData = new SparseCellArray();

                //Attach the CellReader handler.
                CellReader cr = new CellReader(cbFirst50Rows.Checked, CellData, cbFormatValues.Checked);
                xls.VirtualCellStartReading += new VirtualCellStartReadingEventHandler(cr.OnStartReading);
                xls.VirtualCellRead += new VirtualCellReadEventHandler(cr.OnCellRead);

                DateTime StartOpen = DateTime.Now;

                //Open the file. As we have a CellReader attached, the cells won't be loaded into memory, they will be passed to the CellReader
                xls.Open(FileName);
                DateTime StartSheetSelect = cr.StartSheetSelect;
                DateTime EndSheetSelect = cr.EndSheetSelect;

                DateTime EndOpen = DateTime.Now;
                statusBar.Text = "Time to open file: " + (StartSheetSelect - StartOpen).ToString() + "     Time to load file and fill grid: " + (EndOpen - EndSheetSelect).ToString();

                //Set up grid.
                GridCaption.Text = FileName;
                if (CellData != null)
                {
                    DisplayGrid.ColumnCount = CellData.ColCount;
                    DisplayGrid.RowCount = CellData.RowCount;
                }
                else
                {
                    DisplayGrid.ColumnCount = 0;
                    DisplayGrid.RowCount = 0;
                }

                for (int i = 0; i < DisplayGrid.ColumnCount; i++)
                {
                    DisplayGrid.Columns[i].Name = TCellAddress.EncodeColumn(i + 1);
                }
            }
            catch
            {
                GridCaption.Text = "Error Loading File";
                CellData = null;
                throw;
            }
        }

        private void DisplayGrid_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            //Show the row number in the grid at the left
            string r = (e.RowIndex + 1).ToString();
            SizeF textSize = e.Graphics.MeasureString(r, DisplayGrid.Font);
            if (DisplayGrid.RowHeadersWidth < (int)(textSize.Width + 20)) DisplayGrid.RowHeadersWidth = (int)(textSize.Width + 20);
            e.Graphics.DrawString(r, DisplayGrid.Font, SystemBrushes.ControlText, e.RowBounds.Left + DisplayGrid.RowHeadersWidth - textSize.Width - 5, e.RowBounds.Location.Y + ((e.RowBounds.Height - textSize.Height) / 2f));
        }

        private void DisplayGrid_CellValueNeeded(object sender, DataGridViewCellValueEventArgs e)
        {
            if (CellData == null)
            {
                e.Value = null;
                return;
            }

            if (e.RowIndex >= CellData.RowCount)
            {
                e.Value = null;
                return;
            }

            e.Value = CellData.GetValue(e.RowIndex + 1, e.ColumnIndex + 1);
        }
    }
}
