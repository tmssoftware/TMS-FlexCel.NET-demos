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


namespace DebuggingReports
{
    /// <summary>
    /// How to debug a report.
    /// </summary>
    public partial class mainForm: System.Windows.Forms.Form
    {
        public mainForm()
        {
            InitializeComponent();
        }


        private void button1_Click(object sender, System.EventArgs e)
        {
            AutoRun();
        }

        private FlexCelReport CreateReport()
        {
            FlexCelReport Result = new FlexCelReport(true);

            Result.SetValue("test", 3);
            Result.SetValue("tagval", 1);
            Result.SetValue("refval", "l");

            //Here we will add a dummy table with some fantasy values
            DataTable dt = new DataTable("testdb");
            dt.Columns.Add("key", typeof(int));
            dt.Columns.Add("data", typeof(string));
            dt.Rows.Add(new object[] { 5, "cat" });
            dt.Rows.Add(new object[] { 6, "dog" });
            Result.AddTable("testdb", dt, TDisposeMode.DisposeAfterRun);

            return Result;
        }

        public void AutoRun()
        {
            using (FlexCelReport Report = CreateReport())
            {
                string DataPath = Path.Combine(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), ".."), "..") + Path.DirectorySeparatorChar;

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    Report.Run(DataPath + "Debugging Reports.template.xls", saveFileDialog1.FileName);

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
    }

}
