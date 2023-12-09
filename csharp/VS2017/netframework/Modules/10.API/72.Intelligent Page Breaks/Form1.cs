using System;
using System.Drawing;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Text;

using FlexCel.Render;

namespace IntelligentPageBreaks
{
    /// <summary>
    /// Demo showing how to create intelligent page breaks with the API.
    /// </summary>
    public partial class mainForm: System.Windows.Forms.Form
    {
        public mainForm()
        {
            InitializeComponent();
        }

        private Dictionary<string, string> Keywords = CreateKeywords();


        private static Dictionary<string, string> CreateKeywords()
        {
            // A very silly syntax highlighter. We don't have any context here, so for example "get" will be highlighted when it is a property or when it is not, but it is ok for this demo.
            Dictionary<string, string> Result = new Dictionary<string, string>();

            Result.Add("private", null);
            Result.Add("public", null);
            Result.Add("protected", null);
            Result.Add("internal", null);
            Result.Add("static", null);
            Result.Add("void", null);
            Result.Add("get", null);
            Result.Add("set", null);
            Result.Add("return", null);
            Result.Add("while", null);
            Result.Add("for", null);
            Result.Add("foreach", null);
            Result.Add("using", null);
            Result.Add("true", null);
            Result.Add("false", null);

            return Result;
        }

        private string PathToExe
        {
            get
            {
                return Path.Combine(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), ".."), "..") + Path.DirectorySeparatorChar;
            }
        }

        private TRichString SyntaxColor(ExcelFile Xls, int NormalFont, int CommentFont, int HighlightFont, string line)
        {
            List<TRTFRun> RTFRunList = new List<TRTFRun>();

            int i = 0;
            while (i < line.Length)
            {
                if (i > 0 && line[i - 1] == '/' && line[i] == '/')
                {
                    TRTFRun rtf;
                    rtf.FirstChar = i - 1;
                    rtf.FontIndex = CommentFont;
                    RTFRunList.Add(rtf);
                    return new TRichString(line, RTFRunList.ToArray(), Xls);

                }

                int start = i;
                while (i < line.Length && char.IsLetterOrDigit(line[i]))
                {
                    i++;
                }

                if (i > start && Keywords.ContainsKey(line.Substring(start, i - start)))
                {
                    TRTFRun rtf;
                    rtf.FirstChar = start;
                    rtf.FontIndex = HighlightFont;
                    RTFRunList.Add(rtf);
                    rtf.FirstChar = i;
                    rtf.FontIndex = NormalFont;
                    RTFRunList.Add(rtf);
                }

                i++;
            }


            return new TRichString(line, RTFRunList.ToArray(), Xls);
        }

        private void DumpFile(ExcelFile Xls, ref int Row)
        {
            TFlxFont fnt = Xls.GetDefaultFont;
            fnt.Color = Color.Blue;
            int HighlightFont = Xls.AddFont(fnt);
            fnt.Color = Color.Green;
            int CommentFont = Xls.AddFont(fnt);

            int Level = 0;
            Stack<int> LevelStart = new Stack<int>();
            LevelStart.Push(Row);

            using (StreamReader sr = new StreamReader(Path.Combine(PathToExe, "Form1.cs")))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    //Find the level of "keep together" for the row. We will use #region and "{" delimiters
                    //to increase the level. If possible, we would want those blocks together in one page.
                    string s = line.Trim();
                    if (s.StartsWith("#region")) { Level++; LevelStart.Push(Row); }
                    if (s == "{")
                    {
                        Level++;
                        LevelStart.Push(Row - 1);//On {} blocks, we want to keep lines together starting with the previous statement.
                    }

                    if (s == "#endregion" || s == "}")
                    {
                        Level--;
                        Xls.KeepRowsTogether(LevelStart.Pop(), Row, Level + 1, false);
                    }

                    Xls.KeepRowsTogether(Row, Row, Level, true);


                    Xls.SetCellValue(Row, 1, SyntaxColor(Xls, 0, CommentFont, HighlightFont, line.Replace("\t", "    ")));
                    Row++;
                }
            }
        }

        private void AddData(ExcelFile Xls)
        {

            //Fill the file with the contents of this c# file, many times so we can see many page breaks.
            int Row = 3;
            DumpFile(Xls, ref Row);

            Xls.AutofitRowsOnWorkbook(false, true, 1);
            Xls.AutoPageBreaks(50, 100); // we will use a 100% of page scale since we are printing to pdf. 
                                         //If this was to create an Excel file, pagescale should be lower to 
                                         //compensate the differences between page sizes in diiferent printers in Excel

            //Export the file to PDF so we can see the page breaks.
            using (FlexCelPdfExport pdf = new FlexCelPdfExport(Xls, true))
            {
                pdf.Export(saveFileDialog1.FileName);
            }

        }


        private void button1_Click(object sender, System.EventArgs e)
        {
            AutoRun();
        }

        public void AutoRun()
        {
            if (saveFileDialog1.ShowDialog() != DialogResult.OK) return;
            ExcelFile Xls = new XlsFile(true);
            Xls.NewFile(1, TExcelFileFormat.v2019);
            Xls.SetColWidth(1, 78 * 256); //;make longer lines wrap in the cell.
            TFlxFormat fmt = Xls.GetFormat(Xls.GetColFormat(1));
            fmt.WrapText = true;

            Xls.SetColFormat(1, Xls.AddFormat(fmt));
            AddData(Xls);
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
