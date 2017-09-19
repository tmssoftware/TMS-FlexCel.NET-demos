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


namespace UserTables
{
    /// <summary>
    /// Using tables that are defined in the template.
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
            using (FlexCelReport genericReport = new FlexCelReport(true))
            {
                genericReport.UserTable += new UserTableEventHandler(genericReport_UserTable);
                genericReport.DeleteEmptyRanges = false;

                string DataPath = Path.Combine(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), ".."), "..") + Path.DirectorySeparatorChar;

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    genericReport.Run(DataPath + "User Tables.template" + Path.GetExtension(saveFileDialog1.FileName), saveFileDialog1.FileName);

                    if (MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        Process.Start(saveFileDialog1.FileName);
                    }
                }
            }
        }

        private void genericReport_UserTable(object sender, UserTableEventArgs e)
        {
            DataSet ds = new DataSet();

            //On this example we will just return the table with the name specified on parameters
            //but you could return whatever you wanted here.
            //As always, remember to *validate* what the user can enter on the parameters string.

            switch (e.Parameters.ToUpper(CultureInfo.InvariantCulture))
            {
                case "SUPPLIERS":
                    SharedData.Fill(ds, "select * from suppliers", e.TableName);
                    break;
                case "CATEGORIES":
                    SharedData.Fill(ds, "select * from categories", e.TableName);
                    break;
                case "PRODUCTS":
                    SharedData.Fill(ds, "select * from products", e.TableName);
                    break;

                default: throw new Exception("Invalid parameter to user table: " + e.Parameters);
            }

            ((FlexCelReport)sender).AddTable(ds, TDisposeMode.DisposeAfterRun);
        }

        private void btnCancel_Click(object sender, System.EventArgs e)
        {
            Close();
        }
    }


}
