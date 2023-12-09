using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Data.OleDb;
using System.Threading;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using FlexCel.Report;


namespace GenericReports2
{
    /// <summary>
    /// A generic report.
    /// </summary>
    public partial class mainForm: System.Windows.Forms.Form
    {
        private EnterSQLDialog SqlDialog;

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

        private void button2_Click(object sender, System.EventArgs e)
        {
            Close();
        }

        private void btnOpenconnection_Click(object sender, System.EventArgs e)
        {
            string DataPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar;
            string ConfigFile = DataPath + "GenericReports2.udl";
            if (!File.Exists(ConfigFile))
                using (FileStream f = File.Create(ConfigFile))
                {
                    //Nothing, create an empty udl.
                }

            using (Process p = new Process())
            {               
                p.StartInfo.FileName = ConfigFile;
                p.StartInfo.UseShellExecute = true;
                p.Start();
            }              
        }

        private void btnQuery_Click(object sender, System.EventArgs e)
        {
            string DataPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar;
            string ConfigFile = DataPath + "GenericReports2.udl";
            Connection.Close();
            dataSet = new DataSet();


            Connection.ConnectionString = "File Name = " + ConfigFile;

            Connection.Open();

            if (SqlDialog == null) SqlDialog = new EnterSQLDialog();

            if (SqlDialog.ShowDialog() != DialogResult.OK)
                return;

            dbDataAdapter.SelectCommand = new OleDbCommand(SqlDialog.SQL, Connection);
            dbDataAdapter.Fill(dataSet, "Table");
            dataGrid.CaptionText = dbDataAdapter.SelectCommand.CommandText;
            dataGrid.SetDataBinding(dataSet, "Table");
        }

        private void Export(string SQL, out string DataPath)
        {
            Report.ClearTables();
            Report.AddTable(dataSet);
            Report.SetValue("Date", DateTime.Now);
            Report.SetValue("ReportCaption", SQL);
            Report.SetUserFunction("datatype", new DataTypeImp());

            DataPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar; //First try to find the template on exe folder.

            if (!File.Exists(DataPath + "Generic Reports 2.template.xls")) //When on design mode, search for the template 2 folders up.
                DataPath = Path.Combine(DataPath, Path.Combine("..", "..")) + Path.DirectorySeparatorChar;
        }

        private void btnExportExcel_Click(object sender, System.EventArgs e)
        {
            string DataPath = null;
            if (dbDataAdapter == null || dbDataAdapter.SelectCommand == null || dbDataAdapter.SelectCommand.CommandText == null)
            {
                MessageBox.Show("You need to select a query first");
                return;
            }
            Export(dbDataAdapter.SelectCommand.CommandText, out DataPath);

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Report.Run(DataPath + "Generic Reports 2.template.xls", saveFileDialog1.FileName);

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

    /// <summary>
    /// A small used-defined function to know the type of the value inserted.
    /// </summary>
    public class DataTypeImp: TFlexCelUserFunction
    {
        public override object Evaluate(object[] parameters)
        {
            if (parameters == null || parameters.Length != 1)
            {
                throw new Exception("DataType must be called with 1 parameter.");
            }

            if (parameters[0] is double) return "double";
            if (parameters[0] is DateTime) return "datetime";
            return "";
        }
    }
}
