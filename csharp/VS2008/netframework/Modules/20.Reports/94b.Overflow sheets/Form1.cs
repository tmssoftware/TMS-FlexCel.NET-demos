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


namespace Overflowsheets
{
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
            using (FlexCelReport Report = new FlexCelReport(true))
            {
                TMyData[] Data = new TMyData[1010];
                for (int i = 0; i < Data.Length; i++)
                {
                    Data[i] = new TMyData("Customer " + i.ToString());
                }
                Report.AddTable("data", Data);
                Report.SetValue("split", 40);

                string DataPath = Path.Combine(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), ".."), "..") + Path.DirectorySeparatorChar;

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    Report.Run(DataPath + "Overflow Sheets.template.xlsx", saveFileDialog1.FileName);

                    if (MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        Process.Start(saveFileDialog1.FileName);
                    }
                }
            }
        }

        private void btnCancel_Click(object sender, System.EventArgs e)
        {
            Close();
        }
    }

    class TMyData
    {
        public string Name { get; set; }

        public TMyData(string name)
        {
            this.Name = name;
        }
    }

}
