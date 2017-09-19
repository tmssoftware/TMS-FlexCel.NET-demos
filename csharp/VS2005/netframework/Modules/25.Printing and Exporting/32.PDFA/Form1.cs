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
using System.Drawing.Drawing2D;
using FlexCel.Pdf;

using FlexCel.Render;


namespace PDFA
{
    /// <summary>
    /// Exporting xls files to PDF/A.
    /// </summary>
    public partial class mainForm: System.Windows.Forms.Form
    {

        public mainForm()
        {
            InitializeComponent();
            cbPdfType.SelectedIndex = 1;
        }

        private void button2_Click(object sender, System.EventArgs e)
        {
            Close();
        }

        private void export_Click(object sender, System.EventArgs e)
        {
            bool EmbedSource = cbEmbed.Checked;
            TPdfType PdfType = GetPdfType();
            TTagMode TagMode = GetTagMode();

            if (EmbedSource)
            {
                if (PdfType != TPdfType.PDFA3 && PdfType != TPdfType.Standard)
                {
                    MessageBox.Show("To embed a file, you need to use standard PDF or PDF/A3");
                    return;
                }
            }

            if (exportDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            CreateFile(exportDialog.FileName, EmbedSource, PdfType, TagMode);

            if (MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                Process.Start(exportDialog.FileName);
            }

        }

        private TPdfType GetPdfType()
        {
            switch (cbPdfType.SelectedIndex)
            {
                case 0: return TPdfType.Standard;
                case 1:
                case 2: return TPdfType.PDFA1;
                case 3:
                case 4: return TPdfType.PDFA2;
                case 5:
                case 6: return TPdfType.PDFA3;
            }

            throw new Exception("Unexpected PDF type");
        }

        private TTagMode GetTagMode()
        {
            switch (cbPdfType.SelectedIndex)
            {
                case 0:
                case 1:
                case 3:
                case 5: return TTagMode.Full;
            }
            return TTagMode.None;
        }

        private void CreateFile(string FileName, bool EmbedSource, TPdfType PdfType, TTagMode TagMode)
        {
            ExcelFile xls = CreateSourceFile();
            using (FlexCelPdfExport pdf = new FlexCelPdfExport(xls, true))
            {
                pdf.PdfType = PdfType;
                pdf.TagMode = TagMode;
                if (EmbedSource)
                {
                    pdf.AttachFile("Report.xlsx", StandardMimeType.Xlsx, "This is the source file used to create the PDF", DateTime.Now, TPdfAttachmentKind.Source,
                       delegate(TPdfAttachmentWriter attachWriter)
                        {
                            using (MemoryStream ms = new MemoryStream())
                            {
                                xls.Save(ms, TFileFormats.Xlsx);
                                ms.Position = 0;
                                attachWriter.Write(ms);
                            }
                        });
                }
                pdf.Export(FileName);
            }
        }

        private ExcelFile CreateSourceFile()
        {
            ExcelFile xls = new XlsFile();
            xls.NewFile(1, TExcelFileFormat.v2010);
            xls.SetCellValue(1, 1, "This is a test from FlexCel!");
            xls.SetCellValue(2, 1, "Here is some emoji to show unicode surrogate support: 🐜🐏");
            xls.SetCellValue(3, 1, "You might need a font able to show emoji for those characters to show");
            xls.SetCellValue(4, 1, "Windows 7 and 8 have SegoeUISymbol, which can show them and is used automatically by FlexCel.");
            return xls;
        }

    }
}
