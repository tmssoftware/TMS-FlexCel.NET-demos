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

namespace ConsolidatingFiles
{
    /// <summary>
    /// A demo on how to copy many sheets from different files into one file.
    /// </summary>
    public partial class mainForm: System.Windows.Forms.Form
    {

        public mainForm()
        {
            InitializeComponent();
        }

        /// <summary>
        /// This is the method that will be called by the ASP.NET front end. It returns an array of bytes 
        /// with the report data, so the ASP.NET application can stream it to the client.
        /// </summary>
        /// <param name="fileDatas"></param>
        /// <param name="fileNames"></param>
        /// <param name="OnlyData"></param>
        /// <returns>The generated file as a byte array.</returns>
        public byte[] WebRun(Stream[] fileDatas, string[] fileNames, bool OnlyData)
        {
            if (fileNames.Length <= 0)
            {
                throw new ApplicationException("You must select at least one file");
            }

            ExcelFile XlsOut = Consolidate(fileDatas, fileNames, OnlyData);

            using (MemoryStream OutStream = new MemoryStream())
            {
                XlsOut.Save(OutStream);
                return OutStream.ToArray();
            }


        }

        private ExcelFile Consolidate(Stream[] fileDatas, string[] fileNames, bool OnlyData)
        {
            ExcelFile XlsIn = new XlsFile();
            ExcelFile XlsOut = new XlsFile(true);
            XlsOut.NewFile(1, TExcelFileFormat.v2019);

            if (fileNames.Length > 1 && cbOnlyData.Checked) XlsOut.InsertAndCopySheets(1, 2, fileNames.Length - 1);

            for (int i = 0; i < fileNames.Length; i++)
            {
                if (fileDatas != null) XlsIn.Open(fileDatas[i]);
                else XlsIn.Open(fileNames[i]);
                XlsIn.ConvertFormulasToValues(true); //If there is any formula referring to other sheet, convert it to value. 
                                                     //We could also call an overloaded version of InsertAndCopySheets() that
                                                     //copies many sheets at the same time, so references are kept.
                XlsOut.ActiveSheet = i + 1;

                if (OnlyData)
                    XlsOut.InsertAndCopyRange(TXlsCellRange.FullRange(), 1, 1, 1, TFlxInsertMode.ShiftRangeDown, TRangeCopyMode.All, XlsIn, 1);
                else
                {
                    XlsOut.InsertAndCopySheets(1, XlsOut.ActiveSheet, 1, XlsIn);
                }

                //Change sheet name.
                string s = Path.GetFileName(fileNames[i]);
                if (s.Length > 32) XlsOut.SheetName = s.Substring(0, 29) + "...";
                else XlsOut.SheetName = s;

            }

            if (!cbOnlyData.Checked)
            {
                XlsOut.ActiveSheet = XlsOut.SheetCount;
                XlsOut.DeleteSheet(1);  //Remove the empty sheet that came with the workbook.
            }

            XlsOut.ActiveSheet = 1;
            return XlsOut;
        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            if (openFileDialog1.ShowDialog() != DialogResult.OK) return;
            string[] fileNames = openFileDialog1.FileNames;
            if (fileNames.Length <= 0)
            {
                MessageBox.Show("You must select at least one file");
                return;
            }

            ExcelFile XlsOut = Consolidate(null, fileNames, cbOnlyData.Checked);

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                XlsOut.Save(saveFileDialog1.FileName);

                if (MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    Process.Start(saveFileDialog1.FileName);
                }
            }

        }
    }
}
