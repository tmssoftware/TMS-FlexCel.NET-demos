using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using System.IO;
using System.Diagnostics;
using System.Reflection;

namespace GettingStarted
{
    /// <summary>
    /// A small example on how to create a simple file with the API.
    /// Note that you can use the APIMate tool (in Start Menu->TMS FlexCel Studio->Tools) to find out the 
    /// methods you need to call.
    /// </summary>
    public partial class mainForm: System.Windows.Forms.Form
    {

        public mainForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            ExcelFile Xls = new XlsFile(true);
            AddData(Xls);

            if (cbAutoOpen.Checked)
                AutoOpen(Xls);
            else
                NormalOpen(Xls);
        }

        private void AddData(ExcelFile Xls)
        {
            //Create a new file. We could also open an existing file with Xls.Open
            Xls.NewFile(1);
            //Set some cell values.
            Xls.SetCellValue(1, 1, "Hello to the world");
            Xls.SetCellValue(2, 1, 3);
            Xls.SetCellValue(3, 1, 2.1);
            Xls.SetCellValue(4, 1, new TFormula("=Sum(A2,A3)"));

            //Load an image from disk.
            string AssemblyPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            using (Image Img = Image.FromFile(AssemblyPath + Path.DirectorySeparatorChar + ".." + Path.DirectorySeparatorChar + ".." + Path.DirectorySeparatorChar + "Test.bmp"))
            {

                //Add a new image on cell E2
                Xls.AddImage(2, 6, Img);
                //Add a new image with custom properties at cell F6
                Xls.AddImage(Img, new TImageProperties(new TClientAnchor(TFlxAnchorType.DontMoveAndDontResize, 2, 10, 6, 10, 100, 100, Xls), ""));
                //Swap the order of the images. it is not really necessary here, we could have loaded them on the inverse order.
                Xls.BringToFront(1);
            }

            //Add a comment on cell a2
            Xls.SetComment(2, 1, "This is 3");

            //Custom Format cells a2 and a3
            TFlxFormat f = Xls.GetDefaultFormat;
            f.Font.Name = "Times New Roman";
            f.Font.Color = Color.Red;
            f.FillPattern.Pattern = TFlxPatternStyle.LightDown;
            f.FillPattern.FgColor = Color.Blue;
            f.FillPattern.BgColor = Color.White;

            //You can call AddFormat as many times as you want, it will never add a format twice.
            //But if you know the format you are going to use, you can get some extra CPU cycles by
            //calling addformat once and saving the result into a variable.
            int XF = Xls.AddFormat(f);

            Xls.SetCellFormat(2, 1, XF);
            Xls.SetCellFormat(3, 1, XF);

            f.Rotation = 45;
            f.FillPattern.Pattern = TFlxPatternStyle.Solid;
            int XF2 = Xls.AddFormat(f);
            //Apply a custom format to all the row.
            Xls.SetRowFormat(1, XF2);

            //Merge cells
            Xls.MergeCells(5, 1, 10, 6);
            //Note how this one merges with the previous range, creating a final range (5,1,15,6)
            Xls.MergeCells(10, 6, 15, 6);


            //Make the page print in landscape or portrait mode
            bool Landscape = false;  //In this example we will just hard-code it to landscape, in a real app, landscape might have any value.
            if (Landscape)
                Xls.PrintOptions &= ~(TPrintOptions.Orientation | TPrintOptions.NoPls);
            else
            {
                //ALWAYS SET NOPLS TO 0 BEFORE CHANGING THE OTHER OPTIONS.
                Xls.PrintOptions &= ~TPrintOptions.NoPls;
                Xls.PrintOptions |= TPrintOptions.Orientation;
            }

        }


        //This is part of an advanced feature (showing the user using a file) , you do not need to use
        //this method on normal places.
        private string GetLockingUser(string FileName)
        {
            try
            {
                XlsFile xerr = new XlsFile();
                xerr.Open(FileName);
                return " - File might be in use by: " + xerr.Protection.WriteAccess;
            }
            catch
            {
                return String.Empty;
            }
        }

        private void NormalOpen(ExcelFile Xls)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    Xls.Save(saveFileDialog1.FileName);
                }
                catch (IOException ex) //This is not really needed, just to show the username of the user locking the file.
                {
                    throw new IOException(ex.Message + GetLockingUser(saveFileDialog1.FileName), ex);
                }

                if (MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    Process.Start(saveFileDialog1.FileName);
                }
            }
        }

        //This method will use a "trick" to create a temporary file and delete it even when it is open on Excel.
        //We will create a "template" (xlt file), and tell Excel to create a new file based on this template.
        //Then we can safely delete the xlt file, since Excel opened a copy.
        private void AutoOpen(ExcelFile Xls)
        {
            string FilePath = Path.GetTempPath();  //GetTempFileName does not allow us to specify the "xlt" extension.
            string FileName = Path.Combine(FilePath, Guid.NewGuid().ToString() + ".xlt");  //xlt is the extension for excel templates.
            try
            {
                using (FileStream OutStream = new FileStream(FileName, FileMode.Create, FileAccess.Write))
                {
                    FileInfo Fi = new FileInfo(FileName);
                    Fi.Attributes = FileAttributes.Temporary;
                    Xls.IsXltTemplate = true; //Make it an xlt template.
                    Xls.Save(OutStream);
                }
                Process.Start(FileName);
            }
            finally
            {
                File.Delete(FileName);  //As it is an xlt file, we can delete it even when it is open on Excel.			
            }
        }

        /// <summary>
        /// This is the method that will be called by the ASP.NET front end. It returns an array of bytes 
        /// with the report data, so the ASP.NET application can stream it to the client.
        /// </summary>
        /// <returns>The generated file as a byte array.</returns>
        public byte[] WebRun()
        {
            ExcelFile Xls = new XlsFile(true);
            AddData(Xls);

            using (MemoryStream OutStream = new MemoryStream())
            {
                Xls.Save(OutStream);
                return OutStream.ToArray();
            }
        }

    }
}
