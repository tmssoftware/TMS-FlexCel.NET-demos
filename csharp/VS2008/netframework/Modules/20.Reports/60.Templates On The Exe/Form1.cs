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


namespace TemplatesOnTheExe
{
    /// <summary>
    /// How to embed the reports in the executable, including "included" reports.
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
                ordersReport.GetInclude += new GetIncludeEventHandler(ordersReport_GetInclude);
                ordersReport.SetValue("Date", DateTime.Now);

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    Assembly a = Assembly.GetExecutingAssembly();
                    using (Stream InStream = a.GetManifestResourceStream("TemplatesOnTheExe.Templates.Templates On The Exe.template.xls"))
                    {
                        using (FileStream OutStream = new FileStream(saveFileDialog1.FileName, FileMode.Create))
                        {
                            ordersReport.Run(InStream, OutStream);
                        }
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

        private void btnCancel_Click(object sender, System.EventArgs e)
        {
            Close();
        }

        private void ordersReport_GetInclude(object sender, FlexCel.Report.GetIncludeEventArgs e)
        {
            Assembly a = Assembly.GetExecutingAssembly();
            using (Stream InStream = a.GetManifestResourceStream("TemplatesOnTheExe.Templates." + e.FileName))
            {
                byte[] data = new byte[InStream.Length];
                InStream.Position = 0;
                InStream.Read(data, 0, data.Length);
                e.IncludeData = data;
            }
        }
    }

}
