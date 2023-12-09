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


namespace UserDefinedFormats
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
            using (FlexCelReport ordersReport = SharedData.CreateReport())
            {
                ordersReport.SetValue("Date", DateTime.Now);
                ordersReport.SetUserFormat("ZipCode", new ZipCodeImp());
                ordersReport.SetUserFormat("ShipFormat", new ShipFormatImp());

                string DataPath = Path.Combine(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), ".."), "..") + Path.DirectorySeparatorChar;

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    ordersReport.Run(DataPath + "User Defined Formats.template.xlsx", saveFileDialog1.FileName);

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

    }

    #region ZipCode Implementation
    class ZipCodeImp: TFlexCelUserFormat
    {
        public ZipCodeImp()
        {
        }

        public override TFlxPartialFormat Evaluate(ExcelFile workbook, TXlsCellRange rangeToFormat, object[] parameters)
        {
            if (parameters == null || parameters.Length != 1)
                throw new ArgumentException("Bad parameter count in call to ZipCode() user-defined format");

            int color;
            //If the zip code is not valid, don't modify the format.
            if (parameters[0] == null || !int.TryParse(Convert.ToString(parameters[0]), out color)) return new TFlxPartialFormat(null, null, false);

            //This code is not supposed to make sense. We will convert the zip code to a color based in the numeric value.
            TFlxFormat fmt = workbook.GetDefaultFormat;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromArgb(color);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;

            fmt.Font.Color = TExcelColor.FromArgb(~color);

            TFlxApplyFormat apply = new TFlxApplyFormat();
            apply.FillPattern.SetAllMembers(true);
            apply.Font.Color = true;
            return new TFlxPartialFormat(fmt, apply, false);
        }
    }
    #endregion

    #region ShipFormat Implementation
    class ShipFormatImp : TFlexCelUserFormat
    {
        public ShipFormatImp()
        {
        }

        public override TFlxPartialFormat Evaluate(ExcelFile workbook, TXlsCellRange rangeToFormat, object[] parameters)
        {
            //Again, this example is not supposed to make sense, only to show how you can code a complex rule.
            //This method will format the rows with a color that depends in the length of the first parameter,
            //and if the second parameter starts with "B" it will make the text red.

            if (parameters == null || parameters.Length != 2)
                throw new ArgumentException("Bad parameter count in call to ShipFormat() user-defined format");

            int len = Convert.ToString(parameters[0]).Length;
            string country = Convert.ToString(parameters[1]);

            Int32 color = 0xFFFFFF - len * 100;
            TFlxFormat fmt = workbook.GetDefaultFormat;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromArgb(color);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;

            TFlxApplyFormat apply = new TFlxApplyFormat();
            apply.FillPattern.SetAllMembers(true);

            if (country.StartsWith("B"))
            {
                fmt.Font.Color = Colors.OrangeRed;
                apply.Font.Color = true; 
            }

            return new TFlxPartialFormat(fmt, apply, false);
        }
    }
    #endregion

}
