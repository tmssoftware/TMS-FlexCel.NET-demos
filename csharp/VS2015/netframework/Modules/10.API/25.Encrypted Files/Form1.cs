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

namespace EncryptedFiles
{
    /// <summary>
    /// Shows how to deal with Encrypted files.
    /// </summary>
    public partial class mainForm: System.Windows.Forms.Form
    {

        public mainForm()
        {
            InitializeComponent();
        }

        private void btnExit_Click(object sender, System.EventArgs e)
        {
            Close();
        }

        //The event that will actually provide the password to open the empty form.
        private void GetPassword(OnPasswordEventArgs e)
        {
            PasswordDialog Pwd = new PasswordDialog();
            e.Password = string.Empty;
            if (Pwd.ShowDialog() != DialogResult.OK) return;
            e.Password = Pwd.Password;
        }

        private void btnGo_Click(object sender, System.EventArgs e)
        {
            // On this demo we will fill data on an existing file with the api, starting with an encrypted file holding the starting formats.

            // Declare some data for the chart.
            string[] Names = { "Dog", "Cat", "Cow", "Horse", "Fish" };
            int[] Quantities = { 123, 200, 150, 0, 180 };

            // Use two folders up to where the exe is to store the data. (Exe is stored at bin\debug)
            string DataPath = Path.Combine(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), ".."), "..") + Path.DirectorySeparatorChar;

            XlsFile xls = new XlsFile(true);

            // We will use the OnPassword event here to show how to 
            // open a file if you don't know a priory if it is encrypted or not.
            // If you already knew the file was encrypted, (as in this case)you could use:
            // xls.Protection.OpenPassword = "42";

            xls.Protection.OnPassword += new OnPasswordEventHandler(GetPassword);
            xls.Open(Path.Combine(DataPath, "EmptyForm.xls"));

            // Insert rows so the chart range grows. On this case we assume the data is at least 2 rows long. If not, we should handle 
            // the case and do a xls.DeleteRange.
            xls.InsertAndCopyRange(new TXlsCellRange(1, 1, 1, 2), 5, 1, Names.Length - 2, TFlxInsertMode.ShiftRangeDown, TRangeCopyMode.None);

            // Fill the data.
            for (int i = 0; i < Names.Length; i++)
            {
                xls.SetCellValue(4 + i, 1, Names[i]);
                xls.SetCellValue(4 + i, 2, Quantities[i]);
            }

            // Set a new password for opening.
            xls.Protection.OpenPassword = "43";
            xls.Protection.SetModifyPassword("43", false, "Ford Prefect");

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                xls.Save(saveFileDialog1.FileName);

                if (MessageBox.Show("Do you want to open the generated file? (Remember password is 43)", "Confirm", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    Process.Start(saveFileDialog1.FileName);
                }
            }
        }

    }
}
