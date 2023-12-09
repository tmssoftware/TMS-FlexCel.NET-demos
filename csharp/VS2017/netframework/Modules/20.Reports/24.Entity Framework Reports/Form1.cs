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
using System.Transactions;
using System.Configuration;


namespace EntityFrameworkReports
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
    public partial class mainForm : System.Windows.Forms.Form
    {

        public mainForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            AutoRun();
        }

        public void AutoRun()
        {
            using (FlexCelReport ordersReport = new FlexCelReport(true))
            {
                string DataPath = Path.Combine(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), ".."), "..") + Path.DirectorySeparatorChar;
                ordersReport.SetValue("Date", DateTime.Now);

                using (northwndEntities Northwind = new northwndEntities())
                {
                    ordersReport.AddTable("Categories", Northwind.Categories);
                    ordersReport.AddTable("Products", Northwind.Products);

                    if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        TransactionOptions transactionOptions = new TransactionOptions();
                        transactionOptions.IsolationLevel = System.Transactions.IsolationLevel.Serializable; //it would be better to sue Snapshot here, but it isn't supported by SQL Sever CE
                        using (TransactionScope transactionScope = new TransactionScope(TransactionScopeOption.Required, transactionOptions))
                        {
                            ordersReport.Run(DataPath + "Entity Framework Reports.template.xls", saveFileDialog1.FileName);
                            transactionScope.Complete();
                        }

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
        }

        private void btnCancel_Click(object sender, System.EventArgs e)
        {
            Close();
        }
    }

}
