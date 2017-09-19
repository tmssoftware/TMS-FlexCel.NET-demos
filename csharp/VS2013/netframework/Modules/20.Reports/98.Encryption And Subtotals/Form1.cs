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


namespace EncryptionAndSubtotals
{
    public partial class mainForm: System.Windows.Forms.Form
    {
        public mainForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            ManualRun();
        }

        public void ManualRun()
        {
            using (FlexCelReport ordersReport = SharedData.CreateReport())
            {
                ordersReport.BeforeReadTemplate += new GenerateEventHandler(ordersReport_BeforeReadTemplate);
                ordersReport.AfterGenerateSheet += new GenerateEventHandler(ordersReport_AfterGenerateSheet);
                ordersReport.AfterGenerateWorkbook += new GenerateEventHandler(ordersReport_AfterGenerateWorkbook);

                ordersReport.SetValue("Date", DateTime.Now);

                string DataPath = Path.Combine(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), ".."), "..") + Path.DirectorySeparatorChar;

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    ordersReport.Run(DataPath + "Encryption And Subtotals.template.xls", saveFileDialog1.FileName);

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

        private void mainForm_Load(object sender, System.EventArgs e)
        {
            encryptionType.SelectedItem = encryptionType.Items[0];
        }


        private void ordersReport_BeforeReadTemplate(object sender, FlexCel.Report.GenerateEventArgs e)
        {
            e.File.Protection.OpenPassword = OpenPassTemplate.Text;
        }

        private void ordersReport_AfterGenerateSheet(object sender, FlexCel.Report.GenerateEventArgs e)
        {
            e.File.Protection.SetSheetProtection(ProtectSheetPass.Text, new TSheetProtectionOptions(ProtectSheet.Checked));
        }

        private void ordersReport_AfterGenerateWorkbook(object sender, FlexCel.Report.GenerateEventArgs e)
        {
            if (encryptionType.SelectedItem == encryptionType.Items[1])
                e.File.Protection.EncryptionType = TEncryptionType.Xor;
            else e.File.Protection.EncryptionType = TEncryptionType.Standard;
            e.File.Protection.OpenPassword = OpenPassGenerated.Text;
            e.File.Protection.SetModifyPassword(ModifyPassGenerated.Text, RecommendReadOnly.Checked, ReservingUser.Text);
            e.File.Protection.SetWorkbookProtection(ProtectWorkbookPass.Text, new TWorkbookProtectionOptions(false, ProtectWorkbook.Checked));
        }
    }

}
