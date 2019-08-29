using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using FlexCel.Core;
using FlexCel.XlsAdapter;

namespace ReadingFiles
{
    /// <summary>
    /// A demo on how to read a file from FlexCel and display the results.
    /// </summary>
    public partial class mainForm: System.Windows.Forms.Form
    {
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

        private void btnOpenFile_Click(object sender, System.EventArgs e)
        {
            if (openFileDialog1.ShowDialog() != DialogResult.OK) return;
            ImportFile(openFileDialog1.FileName, btnFormatValues.Checked);
        }

        private void ImportFile(string FileName, bool Formatted)
        {
            try
            {
                //Open the Excel file.
                XlsFile xls = new XlsFile(false);
                DateTime StartOpen = DateTime.Now;
                xls.Open(FileName);
                DateTime EndOpen = DateTime.Now;

                //Set up the Grid
                DisplayGrid.DataBindings.Clear();
                DisplayGrid.DataSource = null;
                DisplayGrid.DataMember = null;
                DataSet dataSet1 = new DataSet();
                sheetCombo.Items.Clear();

                //We will create a DataTable "SheetN" for each sheet on the Excel sheet.
                for (int sheet = 1; sheet <= xls.SheetCount; sheet++)
                {
                    xls.ActiveSheet = sheet;

                    sheetCombo.Items.Add(xls.SheetName);

                    DataTable Data = dataSet1.Tables.Add("Sheet" + sheet.ToString());
                    Data.BeginLoadData();
                    try
                    {
                        int ColCount = xls.ColCount;
                        //Add one column on the dataset for each used column on Excel.
                        for (int c = 1; c <= ColCount; c++)
                        {
                            Data.Columns.Add(TCellAddress.EncodeColumn(c), typeof(String));  //Here we will add all strings, since we do not know what we are waiting for.
                        }

                        string[] dr = new string[ColCount];

                        int RowCount = xls.RowCount;
                        for (int r = 1; r <= RowCount; r++)
                        {
                            Array.Clear(dr, 0, dr.Length);
                            //This loop will only loop on used cells. It is more efficient than looping on all the columns.
                            for (int cIndex = xls.ColCountInRow(r); cIndex > 0; cIndex--)  //reverse the loop to avoid calling ColCountInRow more than once.
                            {
                                int Col = xls.ColFromIndex(r, cIndex);

                                if (Formatted)
                                {
                                    TRichString rs = xls.GetStringFromCell(r, Col);
                                    dr[Col - 1] = rs.Value;
                                }
                                else
                                {
                                    int XF = 0; //This is the cell format, we will not use it here.
                                    object val = xls.GetCellValueIndexed(r, cIndex, ref XF);

                                    TFormula Fmla = val as TFormula;
                                    if (Fmla != null)
                                    {
                                        //When we have formulas, we want to write the formula result. 
                                        //If we wanted the formula text, we would not need this part.
                                        dr[Col - 1] = Convert.ToString(Fmla.Result);
                                    }
                                    else
                                    {
                                        dr[Col - 1] = Convert.ToString(val);
                                    }
                                }
                            }
                            Data.Rows.Add(dr);
                        }
                    }
                    finally
                    {
                        Data.EndLoadData();
                    }

                    DateTime EndFill = DateTime.Now;
                    statusBar.Text = String.Format("Time to load file: {0}    Time to fill dataset: {1}     Total time: {2}", (EndOpen - StartOpen).ToString(), (EndFill - EndOpen).ToString(), (EndFill - StartOpen).ToString());

                }

                //Set up grid.
                DisplayGrid.DataSource = dataSet1;
                DisplayGrid.DataMember = "Sheet1";
                sheetCombo.SelectedIndex = 0;
                DisplayGrid.CaptionText = FileName;

            }
            catch
            {
                DisplayGrid.CaptionText = "Error Loading File";
                DisplayGrid.DataSource = null;
                DisplayGrid.DataMember = "";
                sheetCombo.Items.Clear();
                throw;
            }
        }

        private void sheetCombo_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            if ((sender as ComboBox).SelectedIndex < 0) return;
            DisplayGrid.DataMember = "Sheet" + ((sender as ComboBox).SelectedIndex + 1).ToString();
        }

        private void AnalizeFile(string FileName, int Row, int Col)
        {
            XlsFile xls = new XlsFile();
            xls.Open(FileName);

            int XF = 0;
            MessageBox.Show("Active sheet is \"" + xls.ActiveSheetByName + "\"");
            object v = xls.GetCellValue(Row, Col, ref XF);

            if (v == null)
            {
                MessageBox.Show("Cell A1 is empty");
                return;
            }

            //Here we have all the kind of objects FlexCel can return.
            switch (Type.GetTypeCode(v.GetType()))
            {
                case TypeCode.Boolean:
                    MessageBox.Show("Cell A1 is a boolean: " + (bool)v);
                    return;
                case TypeCode.Double:  //Remember, dates are doubles with date format.
                    TUIColor CellColor = Color.Empty;
                    bool HasDate, HasTime;
                    String CellValue = TFlxNumberFormat.FormatValue(v, xls.GetFormat(XF).Format, ref CellColor, xls, out HasDate, out HasTime).ToString();

                    if (HasDate || HasTime)
                    {
                        MessageBox.Show("Cell A1 is a DateTime value: " + FlxDateTime.FromOADate((double)v, xls.OptionsDates1904).ToString() + "\n" +
                            "The value is displayed as: " + CellValue);
                    }
                    else
                    {
                        MessageBox.Show("Cell A1 is a double: " + (double)v + "\n" +
                            "The value is displayed as: " + CellValue + "\n");
                    }
                    return;
                case TypeCode.String:
                    MessageBox.Show("Cell A1 is a string: " + v.ToString());
                    return;
            }

            TFormula Fmla = v as TFormula;
            if (Fmla != null)
            {
                MessageBox.Show("Cell A1 is a formula: " + Fmla.Text + "   Value: " + Convert.ToString(Fmla.Result));
                return;
            }

            TRichString RSt = v as TRichString;
            if (RSt != null)
            {
                MessageBox.Show("Cell A1 is a formatted string: " + RSt.Value);
                return;
            }

            if (v is TFlxFormulaErrorValue)
            {
                MessageBox.Show("Cell A1 is an error: " + TFormulaMessages.ErrString((TFlxFormulaErrorValue)v));
                return;
            }

            throw new Exception("Unexpected value on cell");

        }

        private void btnInfo_Click(object sender, System.EventArgs e)
        {
            MessageBox.Show("This demo shows how to read the contents of an xls file\n" +
                "The 'Open File' button will load an Excel file into a dataset. Depending on the button 'Format Values' it will load the actual values (this is the fastest) or the formatted values.\n" +
                "The 'Format Values' button will modify how the files are read when you press 'Open File'. Formated values are slower, but they will look just how Excel shows them.\n" +
                "The 'Value in Cell A1' button will load an Excel file and show the contents of cell a1 on the active sheet.");
        }

        /// <summary>
        /// This method will not do anything truly useful, but it alows you to see how to 
        /// process the different types of objects that GetCellValue can return
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnValueInCurrentCell_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() != DialogResult.OK) return;
            AnalizeFile(openFileDialog1.FileName, 1, 1);
        }
    }
}
