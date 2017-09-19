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


namespace VirtualDatasets
{
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
            string DataPath = Path.Combine(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), ".."), "..") + Path.DirectorySeparatorChar;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                object[][] SimpleData = LoadDataSet(Path.Combine(DataPath, "Countries.txt"));
                SimpleVirtualArrayDataSource SimpleTable = new SimpleVirtualArrayDataSource(null, SimpleData, new string[] { "Rank", "Country", "Area", "Date" }, "SimpleTable");

                using (FlexCelReport genericReport = new FlexCelReport(true))
                {
                    genericReport.AddTable("SimpleData", SimpleTable);

                    object[][] Complex1 = LoadDataSet(Path.Combine(DataPath, "Countries.txt"));
                    ComplexVirtualArrayDataSource ComplexAreas = new ComplexVirtualArrayDataSource(null, Complex1, new string[] { "Rank", "Country", "Area", "Date" }, "ComplexAreas");
                    object[][] Complex2 = LoadDataSet(Path.Combine(DataPath, "Populations.txt"));
                    ComplexVirtualArrayDataSource ComplexPopulations = new ComplexVirtualArrayDataSource(null, Complex2, new string[] { "Rank", "Country", "Population", "Date" }, "ComplexPopulations");

                    genericReport.AddTable("ComplexAreas", ComplexAreas, TDisposeMode.DisposeAfterRun);
                    genericReport.AddTable("ComplexPopulations", ComplexPopulations, TDisposeMode.DisposeAfterRun);



                    genericReport.Run(Path.Combine(DataPath, "Virtual Datasets.template.xls"), saveFileDialog1.FileName);
                }

                if (MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    Process.Start(saveFileDialog1.FileName);
                }
            }
        }


        private void btnCancel_Click(object sender, System.EventArgs e)
        {
            Close();
        }

        private object[][] LoadDataSet(string filename)
        {
            //Let's create some bussiness object with random data.

            ArrayList Result = new ArrayList();
            using (StreamReader sr = new StreamReader(Path.GetFullPath(filename)))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    string[] fields = line.Split('\t');
                    //Zero validation here since this is a demo and will use always the same data. On a real app you should not expect your data to play nice
                    object[] f = new object[fields.Length];
                    string s = fields[0] as String;
                    f[0] = Convert.ToInt64(s);
                    f[1] = fields[1];
                    s = fields[2] as String;
                    f[2] = (object)Convert.ToInt64(s.Replace(",", ""));
                    f[3] = fields[3];
                    Result.Add(f);
                }
            }

            return (object[][])Result.ToArray(typeof(object[]));
        }
    }

}
