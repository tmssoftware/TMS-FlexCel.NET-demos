using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Threading;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using FlexCel.Report;


namespace GettingStartedReports
{
    /// <summary>
    /// Simple report
    /// </summary>
    public partial class mainForm: System.Windows.Forms.Form
    {
        public mainForm()
        {
            InitializeComponent();
        }

        private void btnGo_Click(object sender, System.EventArgs e)
        {
            //Note that we are using a FlexCelReport component in a form here. We could also create the FlexCelReport component dynamically.

            if (cbAutoOpen.Checked)
                AutoOpenRun();
            else
                NormalRun();
        }

        private void Setup(string UserName, string UserUrl, string DataPath)
        {
            //Set report variables, including an image.

            reportStart.SetValue("Date", DateTime.Now);
            reportStart.SetValue("Name", UserName);
            reportStart.SetValue("TwoLines", "First line" + Environment.NewLine + "Second Line");
            reportStart.SetValue("Empty", null);
            reportStart.SetValue("LinkPage", UserUrl);
            reportStart.SetValue("Img", File.ReadAllBytes(Path.Combine(DataPath, "img.png")));
        }


        private void NormalRun()
        {
            string DataPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            DataPath = Path.Combine(DataPath, "..");
            DataPath = Path.Combine(DataPath, "..");
            Setup(edName.Text, edUrl.Text, DataPath);

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //FlexCel isn't a conversion tool. While it does a good job converting a lot of stuff
                //between xls and xlsx, for best results we will use an xlsx template if the user choose xlsx and xls if the user choose xls.
                reportStart.Run(Path.Combine(DataPath, "Getting Started Reports.template" + Path.GetExtension(saveFileDialog1.FileName)), saveFileDialog1.FileName);

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



        private void AutoOpenRun()
        {
            string DataPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            DataPath = Path.Combine(DataPath, "..");
            DataPath = Path.Combine(DataPath, "..");
            Setup(edName.Text, edUrl.Text, DataPath);

            XlsFile Xls = new XlsFile();
            Xls.Open(Path.Combine(DataPath, "Getting Started Reports.template.xlsx"));
            reportStart.Run(Xls);

            string FilePath = Path.GetTempPath();  //GetTempFileName does not allow us to specify the "xltx" extension.
            string FileName = Path.Combine(FilePath, Guid.NewGuid().ToString() + ".xltx");  //xltx is the extension for excel templates.
            try
            {
                using (FileStream OutStream = new FileStream(FileName, FileMode.Create, FileAccess.ReadWrite))
                {
                    Xls.IsXltTemplate = true; //Make it an xltx template.
                    Xls.Save(OutStream);
                }
                using (Process p = new Process())
                {               
                    p.StartInfo.FileName = FileName;
                    p.StartInfo.UseShellExecute = true;
                    p.Start();
                }              
            }
            finally
            {
                //For .Net 4 and newer you can use Task.Run here. See https://doc.tmssoftware.com/flexcel/net/tips/automatically-open-generated-excel-files.html
                new Thread(delegate()
                {
                    Thread.Sleep(30000); //wait for 30 secs to give Excel time to start.
                    File.Delete(FileName);  //As it is an xltx file, we can delete it even when it is open on Excel.         
                });         
            }
        }

        /// <summary>
        /// This is the method that will be called by the ASP.NET front end. It returns an array of bytes 
        /// with the report data, so the ASP.NET application can stream it to the client.
        /// </summary>
        /// <param name="UserName"></param>
        /// <param name="UserUrl"></param>
        /// <returns>The generated file as a byte array.</returns>
        public byte[] WebRun(string UserName, string UserUrl)
        {
            string DataPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            DataPath = Path.Combine(DataPath, "..");
            DataPath = Path.Combine(DataPath, "..");
            Setup(UserName, UserUrl, DataPath);

            using (MemoryStream OutStream = new MemoryStream())
            {
                using (FileStream InStream = new FileStream(Path.Combine(DataPath, "Getting Started Reports.template.xls"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    reportStart.Run(InStream, OutStream);
                    return OutStream.ToArray();
                }
            }
        }


        private void btnCancel_Click(object sender, System.EventArgs e)
        {
            Close();
        }

    }

}
