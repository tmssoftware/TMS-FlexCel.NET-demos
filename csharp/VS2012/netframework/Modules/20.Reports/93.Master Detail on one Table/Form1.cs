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
using FlexCel.Core;
using FlexCel.XlsAdapter;
using FlexCel.Report;
using FlexCel.Demo.SharedData;


namespace MasterDetailononeTable
{
    /// <summary>
    /// How to split a table into 2.
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

        public void AutoRun()
        {
            using (FlexCelReport ordersReport = SharedData.CreateReport())
            {
                ordersReport.SetValue("Date", DateTime.Now);
                ordersReport.SetValue("ReportCaption", "Sales by year and country");

                using (DataSet ds = new DataSet())
                {
                    SharedData.Fill(ds, @"SELECT Employees.Country, SUM([Order Details].UnitPrice * [Order Details].Quantity) AS Sales, COUNT([Order Details].Quantity) AS OrderCount, DatePart(yyyy, Orders.OrderDate) AS SaleYear, DatePart(q, Orders.OrderDate) AS Quarter FROM ((Employees INNER JOIN Orders ON Employees.EmployeeID = Orders.EmployeeID) INNER JOIN [Order Details] ON Orders.OrderID = [Order Details].OrderID) GROUP BY Employees.Country, DatePart(yyyy, Orders.OrderDate), DatePart(q, Orders.OrderDate)", "Data");
                    ordersReport.AddTable(ds);
                    string DataPath = Path.Combine(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), ".."), "..") + Path.DirectorySeparatorChar;

                    if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        ordersReport.Run(DataPath + "Master Detail on one Table.template.xls", saveFileDialog1.FileName);

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
