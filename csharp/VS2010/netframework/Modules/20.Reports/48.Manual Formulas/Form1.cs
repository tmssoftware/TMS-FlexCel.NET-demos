using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using FlexCel.Report;


namespace ManualFormulas
{
    /// <summary>
    /// Shows the formula tag.
    /// </summary>
    public partial class mainForm: System.Windows.Forms.Form
    {
        public mainForm()
        {
            InitializeComponent();
        }

        private void SetupMines(FlexCelReport MinesReport)
        {
            DataSet ds = new DataSet();
            DataTable dtrows = ds.Tables.Add("datarow");
            dtrows.Columns.Add("position", typeof(int));

            DataTable dtcols = ds.Tables.Add("datacol");
            dtcols.Columns.Add("position", typeof(int));
            dtcols.Columns.Add("value", typeof(int));

            ds.Relations.Add(dtrows.Columns["position"], dtcols.Columns["position"]);

            //let's create 10 mines.
            ArrayList mines = new ArrayList();
            Random rnd = new Random();
            while (mines.Count < 10)
            {
                int nextMine = rnd.Next(9 * 9 - 1);
                int minepos = mines.BinarySearch(nextMine);
                if (minepos >= 0) continue; //the value already exists
                mines.Insert(~minepos, nextMine);
            }

            //Fill the tables on master detail
            for (int r = 0; r < 9; r++)
            {
                dtrows.Rows.Add(new object[] { r });
                for (int c = 0; c < 9; c++)
                {
                    object[] values = new object[2];
                    values[0] = r;
                    if (mines.BinarySearch(r * 9 + c) >= 0) values[1] = 1; else values[1] = DBNull.Value;
                    dtcols.Rows.Add(values);
                }
            }

            //finally, add the tables to the report.
            MinesReport.ClearTables();
            MinesReport.AddTable(ds, TDisposeMode.DisposeAfterRun); //leave to Flexcel to delete the dataset.
        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            AutoRun();
        }

        public void AutoRun()
        {
            using (FlexCelReport MinesReport = new FlexCelReport(true))
            {
                MinesReport.AfterGenerateWorkbook += new GenerateEventHandler(MinesReport_AfterGenerateWorkbook);
                SetupMines(MinesReport);
                string DataPath = Path.Combine(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), ".."), "..") + Path.DirectorySeparatorChar;

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    MinesReport.Run(DataPath + "Manual Formulas.template.xls", saveFileDialog1.FileName);

                    if (MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        using (Process p = new Process())
                        {               
                            p.StartInfo.FileName = saveFileDialog1.FileName;
                            p.StartInfo.UseShellExecute = true;
                            p.Start();
                        }
                    }
                }
            }
        }

        private void btnCancel_Click(object sender, System.EventArgs e)
        {
            Close();
        }

        private void MinesReport_AfterGenerateWorkbook(object sender, FlexCel.Report.GenerateEventArgs e)
        {
            //do some "pretty" up for the final user.
            //we could do this directly on the template, but doing it here allows us to keep the template unprotected and easier to modify.

            e.File.ActiveSheet = 2;
            e.File.SheetVisible = TXlsSheetVisible.Hidden;
            e.File.ActiveSheet = 1;
            e.File.Protection.SetSheetProtection(null, new TSheetProtectionOptions(true));
            for (int r = 20; r <= FlxConsts.Max_Rows97_2003 + 1; r++) e.File.SetRowHidden(r, true);
            for (int c = 12; c <= FlxConsts.Max_Columns97_2003 + 1; c++) e.File.SetColHidden(c, true);
        }
    }

}
