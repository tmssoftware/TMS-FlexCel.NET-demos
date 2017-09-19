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
using FlexCel.Demo.SharedData;


namespace Images
{
    /// <summary>
    /// A report with lots of images.
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
                ordersReport.GetImageData += new GetImageDataEventHandler(ordersReport_GetImageData);
                ordersReport.SetValue("Date", DateTime.Now);

                string DataPath = Path.Combine(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), ".."), "..") + Path.DirectorySeparatorChar;

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    ordersReport.Run(DataPath + "Images.template.xls", saveFileDialog1.FileName);

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

        private void ordersReport_GetImageData(object sender, FlexCel.Report.GetImageDataEventArgs e)
        {
            if (String.Compare(e.ImageName, "<#PhotoCode>", true) == 0)
            {
                byte[] RealImageData = ImageUtils.StripOLEHeader(e.ImageData); //On access databases, images are stored with an OLE 
                //header that we have to strip to get the real image.
                //This is done automatically by flexcel in most cases,
                //but here we have the original image format.
                using (MemoryStream MemStream = new MemoryStream(RealImageData)) //Keep stream open until bitmap has been used
                {
                    using (Bitmap bmp = new Bitmap(MemStream))
                    {
                        bmp.RotateFlip(RotateFlipType.Rotate90FlipNone);
                        using (MemoryStream OutStream = new MemoryStream())
                        {
                            bmp.Save(OutStream, System.Drawing.Imaging.ImageFormat.Png);
                            e.Width = bmp.Width;
                            e.Height = bmp.Height;
                            e.ImageData = OutStream.ToArray();
                        }
                    }
                }
            }

        }
    }

}
