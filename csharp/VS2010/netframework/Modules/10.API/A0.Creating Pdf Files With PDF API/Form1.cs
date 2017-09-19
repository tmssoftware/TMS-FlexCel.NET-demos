using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using FlexCel.Core;
using FlexCel.Pdf;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Drawing.Drawing2D;

namespace CreatingPdfFilesWithPDFAPI
{
    /// <summary>
    /// Jow to create PDF files directly with FlexCel.
    /// </summary>
    public partial class mainForm: System.Windows.Forms.Form
    {
        public mainForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() != DialogResult.OK) return;

            TUITextDecoration Underline = new TUITextDecoration(TUIUnderline.Single);
            PdfWriter pdf = new PdfWriter();

            using (FileStream file = new FileStream(saveFileDialog1.FileName, FileMode.Create))
            {
                pdf.Compress = true;
                pdf.BeginDoc(file);
                pdf.YAxisGrowsDown = true; //To keep it compatible with GDI+
                using (TUIFont f = TUIFont.Create("times new roman", (float)22.5, TUIFontStyle.Italic))
                {
                    using (TUIFont f2 = TUIFont.Create("Arial", (float)12, TUIFontStyle.Italic))
                    {
                        pdf.DrawString("This is the first line on a test of many lines.", f, Underline, Brushes.Navy, 100, 100);
                        pdf.DrawString("Some unicode: \u0e2a\u0e27\u0e31\u0e2a\u0e14\u0e35", f, Underline, Brushes.ForestGreen, 100, 200);
                        pdf.DrawString("More lines here!", f, Underline, Brushes.ForestGreen, 200, 300);
                        pdf.DrawString("And this is the last line.", f, Underline, Brushes.Black, 200, 400);
                        pdf.Properties.Author = "Adrian";
                        pdf.Properties.Title = "This is a test of FlexCel Api";
                        pdf.Properties.Keywords = "test\nflexcel\napi";
                        pdf.NewPage();
                        pdf.SaveState();
                        pdf.Rotate(200, 100, 45);
                        pdf.DrawString("Some rotated test", f, Underline, Brushes.Black, 200, 200);
                        pdf.RestoreState();
                        pdf.DrawString("Some NOT rotated text", f, Underline, Brushes.Black, 200, 200);
                        pdf.DrawString("Hello from FlexCel!", f2, Brushes.Black, 200, 50);

                        TPointF[] points = { new TPointF(200, 100), new TPointF(200, 50), new TPointF(500, 50), new TPointF(700, 100) };
                        pdf.DrawLines(Pens.DarkOrchid, points);

                        RectangleF Coords = new RectangleF(100, 300, 100, 100);
                        using (Brush Gradient = new LinearGradientBrush(Coords, Color.Red, Color.Blue, 0f))
                        {
                            pdf.DrawAndFillRectangle(Pens.Red, Gradient, 100, 300, 100, 100);
                        }
                        pdf.DrawRectangle(Pens.DarkSlateBlue, 100, 300, 50, 50);
                        pdf.DrawLine(Pens.Black, 100, 300, 200, 400);

                        string AssemblyPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                        using (Image Img = Image.FromFile(AssemblyPath + Path.DirectorySeparatorChar + ".." + Path.DirectorySeparatorChar + ".." + Path.DirectorySeparatorChar + "test.jpg"))
                        {
                            pdf.DrawImage(Img, new RectangleF(200, 300, 200, 150), null);
                        }
                        pdf.IntersectClipRegion(new RectangleF(100, 100, 50, 50));
                        pdf.FillRectangle(Brushes.DarkTurquoise, 100, 100, 100, 100);

                        pdf.EndDoc();
                    }
                }
            }
            if (MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                Process.Start(saveFileDialog1.FileName);
            }

        }
    }
}
