using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Resources;
using System.Globalization;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using FlexCel.Report;
using FlexCel.Demo.SharedData;


namespace DirectSQL
{
    /// <summary>
    /// Summary description for Form1.
    /// </summary>
    public partial class mainForm: System.Windows.Forms.Form
    {

        public mainForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            using (FlexCelReport genericReport = new FlexCelReport(true))
            {
                IDbDataAdapter genericAdapter = SharedData.GetDataAdapter();
                try
                {
                    genericReport.SetValue("ReportCaption", "Sales by Country and Employee");
                    genericReport.AddConnection("Northwind", genericAdapter, CultureInfo.CurrentCulture);

                    //In OleDb the parameters are positional, you don't really need to name them when creating them.
                    //But when you are using an SQL Server connection, you *need*
                    //to specify the parameter name ("@StartDate") and make it equal to "@" + the name
                    //of the parameter. It is recommended that you always specify the name, even in OleDb connections.

                    //Also, we are not going to create the parameters directly here (using new SqlCeParameter(...).
                    //We are going to centralize all data access for the demos in SharedData, so we can change it and change all demos.
                    genericReport.AddSqlParameter("StartDate", SharedData.CreateParameter("@StartDate", startDate.Value.Date));
                    genericReport.AddSqlParameter("EndDate", SharedData.CreateParameter("@EndDate", endDate.Value.Date));
                    string DataPath = Path.Combine(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), ".."), "..") + Path.DirectorySeparatorChar;

                    if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        genericReport.Run(DataPath + "Direct SQL.template.xls", saveFileDialog1.FileName);

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
                finally
                {
                    ((IDisposable)genericAdapter).Dispose();
                }
            }
        }

        private void btnCancel_Click(object sender, System.EventArgs e)
        {
            Close();
        }
    }

}
