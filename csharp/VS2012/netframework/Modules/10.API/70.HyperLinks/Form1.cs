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

namespace HyperLinks
{
    /// <summary>
    /// How to deal with Hyperlinks in FlexCel.
    /// </summary>
    public partial class mainForm: System.Windows.Forms.Form
    {

        public mainForm()
        {
            InitializeComponent();
            ResizeToolbar(mainToolbar);
        }

        private void ResizeToolbar(ToolStrip toolbar)
        {

            using (Graphics gr = CreateGraphics())
            {
                double xFactor = gr.DpiX / 96.0;
                double yFactor = gr.DpiY / 96.0;
                toolbar.ImageScalingSize = new Size((int)(24 * xFactor), (int)(24 * yFactor));
                toolbar.Width = 0; //force a recalc of the buttons.
            }

        }

        private void button2_Click(object sender, System.EventArgs e)
        {
            Close();
        }

        private ExcelFile Xls = null;

        private void ReadHyperLinks_Click(object sender, System.EventArgs e)
        {
            if (openFileDialog1.ShowDialog() != DialogResult.OK) return;
            Xls = new XlsFile();

            Xls.Open(openFileDialog1.FileName);

            dataGrid.CaptionText = "Hyperlinks on file: " + openFileDialog1.FileName;
            HlDataTable.Rows.Clear();


            for (int i = 1; i <= Xls.HyperLinkCount; i++)
            {
                TXlsCellRange Range = Xls.GetHyperLinkCellRange(i);
                THyperLink HLink = Xls.GetHyperLink(i);

                string HLinkType = Enum.GetName(typeof(THyperLinkType), HLink.LinkType);

                object[] values ={i, TCellAddress.EncodeColumn(Range.Left)+Range.Top.ToString(),
                                     TCellAddress.EncodeColumn(Range.Right)+Range.Bottom.ToString(),
                                     HLinkType,
                                     HLink.Text,
                                     HLink.Description,
                                     HLink.TextMark,
                                     HLink.TargetFrame,
                                     HLink.Hint
                                };
                HlDataTable.Rows.Add(values);

            }

        }

        private void writeHyperLinks_Click(object sender, System.EventArgs e)
        {
            if (Xls == null)
            {
                MessageBox.Show("You need to open a file first.");
                return;
            }

            ExcelFile XlsOut = new XlsFile(true);
            XlsOut.NewFile(1, TExcelFileFormat.v2019);

            for (int i = 1; i <= Xls.HyperLinkCount; i++)
            {
                TXlsCellRange Range = Xls.GetHyperLinkCellRange(i);
                THyperLink HLink = Xls.GetHyperLink(i);

                int XF = -1;
                object Value = Xls.GetCellValue(Range.Top, Range.Left, ref XF);
                XlsOut.SetCellValue(i, 1, Value, XlsOut.AddFormat(Xls.GetFormat(XF)));
                XlsOut.AddHyperLink(new TXlsCellRange(i, 1, i, 1), HLink);
            }

            if (saveFileDialog1.ShowDialog() != DialogResult.OK) return;
            XlsOut.Save(saveFileDialog1.FileName);
            if (MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                Process.Start(saveFileDialog1.FileName);
            }
        }

    }
}
