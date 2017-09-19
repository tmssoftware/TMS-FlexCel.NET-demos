using System;
using System.Drawing;
using System.Collections;
using System.Windows.Forms;
using System.Data;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Text;

namespace AdvancedAPI
{
    /// <summary>
    /// A demo on creating a file using more advanced features.
    /// </summary>
    public partial class mainForm: System.Windows.Forms.Form
    {

        public mainForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            ExcelFile Xls = new XlsFile(true);
            AddData(Xls);

            NormalOpen(Xls);
        }

        /// <summary>
        /// We will use this path to find the template.xls. Code is a little complex because it has to run in mono.
        /// </summary>
		private string PathToExe
        {
            get
            {
                return Path.Combine(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), ".."), "..") + Path.DirectorySeparatorChar;
            }
        }

        //some silly data to fill in the cells. A real app would read this from somewhere else.
        string[] Country = { "USA", "Canada", "Spain", "France", "United Kingdom", "Australia", "Brazil", "Unknown" };

        int DataRows = 100;

        /// <summary>
        /// Will return a list of countries separated by Character(0) so it can be used as input for a built in list.
        /// </summary>
        /// <returns></returns>
        private string GetCountryList()
        {
            StringBuilder sb = new StringBuilder();
            string sep = "";
            foreach (string c in Country)
            {
                sb.Append(sep);
                sb.Append(c);
                sep = "\0";  //not very efficient method to concat, but good enough for this demo.
            }

            return sb.ToString();
        }

        private void AddChart(TXlsNamedRange DataCell, ExcelFile Xls)
        {
            //Find the cell where the cart will go.
            TXlsNamedRange ChartRange = Xls.GetNamedRange("ChartData", -1);

            //Insert cells to expand the range for the chart. It already has 2 rows, so we need to insert Country.Length - 2
            //Note also that we insert after ChartRange.Top, so the chart is updates with the new range.
            Xls.InsertAndCopyRange(new TXlsCellRange(ChartRange.Top, ChartRange.Left, ChartRange.Top, ChartRange.Left + 1),
                ChartRange.Top + 1, ChartRange.Left, Country.Length - 2, TFlxInsertMode.ShiftRangeDown);  //we use shiftrangedown so not all the row goes down and the chart stays in place.

            //Get the cell addresses of the data range.
            TCellAddress FirstCell = new TCellAddress(DataCell.Top, DataCell.Left);
            TCellAddress SecondCell = new TCellAddress(DataCell.Top + DataRows, DataCell.Left + 1);
            TCellAddress FirstSumCell = new TCellAddress(DataCell.Top, DataCell.Left + 1);

            //Fill a table with the data to be used in the chart.
            for (int r = ChartRange.Top; r < ChartRange.Top + Country.Length; r++)
            {
                Xls.SetCellValue(r, ChartRange.Left, Country[r - ChartRange.Top]);
                Xls.SetCellValue(r, ChartRange.Left + 1, new TFormula("=SUMIF(" + FirstCell.CellRef + ":" + SecondCell.CellRef +
                    ",\"" + Country[r - ChartRange.Top] + "\", " + FirstSumCell.CellRef + ":" + SecondCell.CellRef + ")"));
            }

        }

        private void AddData(ExcelFile Xls)
        {
            string TemplateFile = "template.xls";
            if (cbXlsxTemplate.Checked)
            {
                if (!XlsFile.SupportsXlsx)
                {
                    throw new Exception("Xlsx files are not supported in this version of the .NET framework");
                }
                TemplateFile = "template.xlsm";
            }

            // Open an existing file to be used as template. In this example this file has
            // little data, in a real situation it should have as much as possible. (Or even better, be a report)
            Xls.Open(Path.Combine(PathToExe, TemplateFile));

            //Find the cell where we want to fill the data. In this case, we have created a named range "data" so the address
            //is not hardcoded here.
            TXlsNamedRange DataCell = Xls.GetNamedRange("Data", -1);

            //Add a chart with totals
            AddChart(DataCell, Xls);
            //Note that "DataCell" will change because we inserted rows above it when creating the chart. But we will keep using the old one.

            //Add the captions. This should probably go into the template, but in a dynamic environment it might go here.
            Xls.SetCellValue(DataCell.Top - 1, DataCell.Left, "Country");
            Xls.SetCellValue(DataCell.Top - 1, DataCell.Left + 1, "Quantity");

            //Add a rectangle around the cells
            TFlxApplyFormat ApplyFormat = new TFlxApplyFormat();
            ApplyFormat.SetAllMembers(false);
            ApplyFormat.Borders.SetAllMembers(true);  //We will only apply the borders to the existing cell formats
            TFlxFormat fmt = Xls.GetDefaultFormat;
            fmt.Borders.Left.Style = TFlxBorderStyle.Double;
            fmt.Borders.Left.Color = Color.Black;
            fmt.Borders.Right.Style = TFlxBorderStyle.Double;
            fmt.Borders.Right.Color = Color.Black;
            fmt.Borders.Top.Style = TFlxBorderStyle.Double;
            fmt.Borders.Top.Color = Color.Black;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Double;
            fmt.Borders.Bottom.Color = Color.Black;
            Xls.SetCellFormat(DataCell.Top - 1, DataCell.Left, DataCell.Top, DataCell.Left + 1, fmt, ApplyFormat, true);  //Set last parameter to true so it draws a box.

            //Freeze panes
            Xls.FreezePanes(new TCellAddress(DataCell.Top, 1));


            Random Rnd = new Random();

            //Fill the data
            int z = 0;
            int OutlineLevel = 0;
            for (int r = 0; r <= DataRows; r++)
            {

                //Fill the values.
                Xls.SetCellValue(DataCell.Top + r, DataCell.Left, Country[z % Country.Length]);  //For non C# users, "%" means "mod" or modulus in other languages. It is the rest of the integer division.
                Xls.SetCellValue(DataCell.Top + r, DataCell.Left + 1, Rnd.Next(1000));

                //Add the country to the outline
                Xls.SetRowOutlineLevel(DataCell.Top + r, OutlineLevel);
                //increment the country randomly
                if (Rnd.Next(3) == 0)
                {
                    z++;
                    OutlineLevel = 0;  //Break the group and create a new one. 
                }
                else
                {
                    OutlineLevel = 1;
                }
            }

            //Make the "+" signs of the outline appear at the top.
            Xls.OutlineSummaryRowsBelowDetail = false;

            //Collapse the outline to the first level.
            Xls.CollapseOutlineRows(1, TCollapseChildrenMode.Collapsed);

            //Add Data Validation for the first column, it must be a country.
            TDataValidationInfo dv = new TDataValidationInfo(
                TDataValidationDataType.List, //We will use a built in list.
                TDataValidationConditionType.Between,  //This parameter does not matter since it is a list. It will not be used.
                "=\"" + GetCountryList() + "\"",   //We could have used a range of cells here with the values (like "=C1..C4") Instead, we directly entered the list in the formula.
                null,  //no need for a second formula, not used in List
                false,
                true,
                true,  //Note that as we entered the data directly in FirstFormula, we need to set this to true
                true,
                "Unknown country",
                "Please make sure that the country is in the list",
                false, //We will not use an input box, so this is false and the 2 next entries are null
                null,
                null,
                TDataValidationIcon.Stop);  //We will use the stop icon so no invalid input is permitted.
            Xls.AddDataValidation(new TXlsCellRange(DataCell.Top, DataCell.Left, DataCell.Top + DataRows, DataCell.Left), dv);

            //Add Data Validation for the second column, it must be an integer between 0 and 1000.
            dv = new TDataValidationInfo(
                TDataValidationDataType.WholeNumber, //We will request a number.
                TDataValidationConditionType.Between,
                "=0",  //First formula marks the first part of the "between" condition.
                "=1000",  //Second formula is the second part.
                false,
                false,
                false,
                true,
                "Invalid Quantity",
                null, //We will leave the default error message.
                true,
                "Quantity:",
                "Please enter a quantity between 0 and 1000",
                TDataValidationIcon.Stop);  //We will use the stop icon so no invalid input is permitted.
            Xls.AddDataValidation(new TXlsCellRange(DataCell.Top, DataCell.Left + 1, DataCell.Top + DataRows, DataCell.Left + 1), dv);


            //Search country "Unknown" and replace it by "no".
            //This does not make any sense here (we could just have entered "no" to begin)
            //but it shows how to do it when modifying an existing file
            Xls.Replace("Unknown", "no", TXlsCellRange.FullRange(), true, false, true);

            //Autofit the rows. As we keep the row height automatic this will not show when opening in Excel, but will work when directly printing from FlexCel.
            Xls.AutofitRowsOnWorkbook(false, true, 1);

            Xls.Recalc(); //Calculate the SUMIF formulas so we can sort by them. Note that FlexCel automatically recalculates before saving,
                          //but in this case we haven't saved yet, so the sheet is not recalculated. You do not normally need to call Recalc directly.

            //Sort the data. As in the case with replace, this does not make much sense. We could have entered the data sorted to begin
            //But it shows how you can use the feature.

            //Find the cell where the chart goes.
            TXlsNamedRange ChartRange = Xls.GetNamedRange("ChartData", -1);
            Xls.Sort(new TXlsCellRange(ChartRange.Top, ChartRange.Left, ChartRange.Top + Country.Length, ChartRange.Left + 1),
                true, new int[] { 2 }, new TSortOrder[] { TSortOrder.Descending }, null);



            //Protect the Sheet
            TSheetProtectionOptions Sp = new TSheetProtectionOptions(false); //Create default protection options that allows everything.
            Sp.InsertColumns = false; //Restrict inserting columns.
            Xls.Protection.SetSheetProtection("flexcel", Sp);
            //Set a modify password. Note that this does *not* encrypt the file.
            Xls.Protection.SetModifyPassword("flexcel", true, "flexcel");

            Xls.Protection.OpenPassword = "flexcel";  //OpenPasword is the only password that will actually encrypt the file, so you will not be able to open it with flexcel if you do not know the password.

            //Select cell A1
            Xls.SelectCell(1, 1, true);
        }

        //This is part of an advanced feature (showing the user using a file) , you do not need to use
        //this method on normal places.
        private string GetLockingUser(string FileName)
        {
            try
            {
                XlsFile xerr = new XlsFile();
                xerr.Open(FileName);
                return " - File might be in use by: " + xerr.Protection.WriteAccess;
            }
            catch
            {
                return String.Empty;
            }
        }

        private void NormalOpen(ExcelFile Xls)
        {
            if (cbXlsxTemplate.Checked) saveFileDialog1.FilterIndex = 1; else saveFileDialog1.FilterIndex = 0;
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (!XlsFile.SupportsXlsx && Path.GetExtension(saveFileDialog1.FileName) == ".xlsm")
                {
                    throw new Exception("Xlsx files are not supported in this version of the .NET framework");
                }


                try
                {
                    Xls.Save(saveFileDialog1.FileName);
                }
                catch (IOException ex) //This is not really needed, just to show the username of the user locking the file.
                {
                    throw new IOException(ex.Message + GetLockingUser(saveFileDialog1.FileName), ex);
                }

                if (MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    Process.Start(saveFileDialog1.FileName);
                }
            }
        }
    }
}
