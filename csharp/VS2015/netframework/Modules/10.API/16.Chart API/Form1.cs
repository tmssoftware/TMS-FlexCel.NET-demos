using System;
using System.Drawing;
using System.Collections;
using System.Windows.Forms;
using System.Data;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Text;

namespace ChartAPI
{
    /// <summary>
    /// A demo on creating a chart with code.
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

        private string PathToExe
        {
            get
            {
                return Path.Combine(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), ".."), "..") + Path.DirectorySeparatorChar;
            }
        }

        private void AutoRun()
        {
            //We will use data already stored in a file. For a real case, you would
            //probably fill this data from some database.
            string fileName = Path.Combine(PathToExe, "git-stats.xlsx");
            ExcelFile Xls = new XlsFile(fileName, true);
            
            //Add a new empty sheet for adding the chart.
            Xls.InsertAndCopySheets(0, 1, 1);
            Xls.ActiveSheet = 1;
            Xls.SheetName = "Chart";
            Xls.PrintToFit = true;
            Xls.PrintScale = 70;
            Xls.PrintXResolution = 600;
            Xls.PrintYResolution = 600;
            Xls.PrintOptions = TPrintOptions.None;
            Xls.PrintPaperSize = TPaperSize.Letter;
            Xls.PrintLandscape = true;

            AddChart(Xls);
            NormalOpen(Xls);
        }



        private void AddChart(ExcelFile xls)
        {
            //This code is adapted from APIMate.
            //Objects
            TShapeProperties ChartOptions1 = new TShapeProperties();
            ChartOptions1.Anchor = new TClientAnchor(TFlxAnchorType.MoveAndResize, 1, 215, 1, 608, 30, 228, 17, 736);
            ChartOptions1.ShapeName = "Lines of code";
            ChartOptions1.Print = true;
            ChartOptions1.Visible = true;
            ChartOptions1.ShapeOptions.SetValue(TShapeOption.fLockText, true);
            ChartOptions1.ShapeOptions.SetValue(TShapeOption.LockRotation, true);
            ChartOptions1.ShapeOptions.SetValue(TShapeOption.fAutoTextMargin, true);
            ChartOptions1.ShapeOptions.SetValue(TShapeOption.fillColor, 134217806);
            ChartOptions1.ShapeOptions.SetValue(TShapeOption.wzName, "Lines of code");
            ExcelChart Chart1 = xls.AddChart(ChartOptions1, TChartType.Area, new ChartStyle(102), false);

            TDataLabel Title = new TDataLabel();
            Title.PositionZeroBased = null;
            ChartFillOptions TextFillOptions = new ChartFillOptions(new TShapeFill(new TSolidFill(TDrawingColor.FromRgb(0x80, 0x80, 0x80)), true, TFormattingType.Subtle, TDrawingColor.FromRgb(0x00, 0x00, 0x00, new TColorTransform(TColorTransformType.Alpha, 0)), false));
            TChartTextOptions LabelTextOptions = new TChartTextOptions(new TFlxChartFont("Calibri Light", 320, TExcelColor.FromArgb(0x80, 0x80, 0x80), TFlxFontStyles.Bold, TFlxUnderline.None, TFontScheme.Major), THFlxAlignment.center, TVFlxAlignment.center, TBackgroundMode.Transparent, TextFillOptions);
            Title.TextOptions = LabelTextOptions;
            TDataLabelOptions LabelOptions = new TDataLabelOptions();
            Title.LabelOptions = LabelOptions;
            ChartLineOptions ChartLineOptions = new ChartLineOptions(new TShapeLine(true, new TLineStyle(new TNoFill(), null), null, TFormattingType.Subtle));
            ChartFillOptions ChartFillOptions = new ChartFillOptions(new TShapeFill(new TNoFill(), false, TFormattingType.Subtle, null, false));
            Title.Frame = new TChartFrameOptions(ChartLineOptions, ChartFillOptions, false);

            TRTFRun[] Runs;
            Runs = new TRTFRun[1];
            Runs[0].FirstChar = 0;
            TFlxFont fnt;
            fnt = xls.GetDefaultFont;
            fnt.Name = "Calibri Light";
            fnt.Size20 = 320;
            fnt.Color = TExcelColor.FromArgb(0x80, 0x80, 0x80);
            fnt.Style = TFlxFontStyles.Bold;
            fnt.Family = 0;
            fnt.CharSet = 1;
            fnt.Scheme = TFontScheme.Major;
            Runs[0].FontIndex = xls.AddFont(fnt);
            TRichString LabelValue1 = new TRichString("FlexCel: Lines of code over time", Runs, xls);

            Title.LabelValues = new Object[] { LabelValue1 };

            Chart1.SetTitle(Title);

            Chart1.Background = new TChartFrameOptions(TDrawingColor.FromTheme(TThemeColor.Dark1, new TColorTransform(TColorTransformType.LumMod, 0.15), new TColorTransform(TColorTransformType.LumOff, 0.85)), 9525, TDrawingColor.FromTheme(TThemeColor.Light1), false);

            TChartFrameOptions PlotAreaFrame;
            ChartLineOptions = new ChartLineOptions(new TShapeLine(true, new TLineStyle(new TNoFill(), null), null, TFormattingType.Subtle));
            ChartFillOptions = new ChartFillOptions(new TShapeFill(new TPatternFill(TDrawingColor.FromTheme(TThemeColor.Dark1, new TColorTransform(TColorTransformType.LumMod, 0.15), new TColorTransform(TColorTransformType.LumOff, 0.85)), TDrawingColor.FromTheme(TThemeColor.Light1), TDrawingPattern.ltDnDiag), true, TFormattingType.Subtle, null, false));
            PlotAreaFrame = new TChartFrameOptions(ChartLineOptions, ChartFillOptions, false);
            TChartPlotAreaPosition PlotAreaPos = new TChartPlotAreaPosition(true, TChartRelativeRectangle.Automatic, TChartLayoutTarget.Inner, true);
            Chart1.PlotArea = new TChartPlotArea(PlotAreaFrame, PlotAreaPos, false);

            Chart1.SetChartOptions(1, new TAreaChartOptions(false, TStackedMode.Stacked, null));

            int LastYear = 0;
            double shade = 1;
            for (int i = 2; i < 190; i++)
            {
                ChartSeries Series = new ChartSeries(
                    "=" + new TCellAddress("Data", 1, i, true, true).CellRef, 
                    "=" + new TCellAddress("Data", 2, i, true, true).CellRef + ":" + new TCellAddress("Data", 189, i, true, true).CellRef, 
                    "=Data!$A$2:$A$189");

                //We will display every year in a single color. Each month gets its own shade.
                int xf = -1;
                int Year = FlxDateTime.FromOADate(((double)xls.GetCellValue(2, 1, i, ref xf)), false).Year;
                if (LastYear != Year) shade = 1; else if (shade > 0.3) shade -= 0.05;
                    LastYear = Year;
                TDrawingColor SeriesColor = TDrawingColor.FromTheme(TThemeColor.Accent1 + Year % 6, 
                    new TColorTransform(TColorTransformType.Shade, shade));

                ChartSeriesFillOptions SeriesFill = new ChartSeriesFillOptions(new TShapeFill(new TSolidFill(SeriesColor), true, TFormattingType.Subtle, null, false), null, false, false);
                ChartSeriesLineOptions SeriesLine = new ChartSeriesLineOptions(new TShapeLine(true, new TLineStyle(new TNoFill(), null), null, TFormattingType.Subtle), false);
                Series.Options.Add(new ChartSeriesOptions(-1, SeriesFill, SeriesLine, null, null, null, true));

                Chart1.AddSeries(Series);
            }

            Chart1.PlotEmptyCells = TPlotEmptyCells.Zero;
            Chart1.ShowDataInHiddenRowsAndCols = false;

            TFlxChartFont AxisFont = new TFlxChartFont("Calibri", 180, TExcelColor.FromArgb(0x59, 0x59, 0x59), TFlxFontStyles.None, TFlxUnderline.None, TFontScheme.Minor);
            TAxisLineOptions AxisLine = new TAxisLineOptions();
            AxisLine.MainAxis = new ChartLineOptions(new TShapeLine(true, new TLineStyle(new TSolidFill(TDrawingColor.FromTheme(TThemeColor.Dark1, new TColorTransform(TColorTransformType.LumMod, 0.15), new TColorTransform(TColorTransformType.LumOff, 0.85))), 9525, TPenAlignment.Center, TLineCap.Flat, TCompoundLineType.Single, null, TLineJoin.Round, null, null, null), null, TFormattingType.Subtle));
            AxisLine.DoNotDrawLabelsIfNotDrawingAxis = false;
            TAxisTickOptions AxisTicks = new TAxisTickOptions(TTickType.Outside, TTickType.None, TAxisLabelPosition.NextToAxis, TBackgroundMode.Transparent, TDrawingColor.FromRgb(0x59, 0x59, 0x59), 0);
            TAxisRangeOptions AxisRangeOptions = new TAxisRangeOptions(12, 1, false, false, false);
            TBaseAxis CatAxis = new TCategoryAxis(0, 0, 12, TDateUnits.Days, 12, TDateUnits.Days, TDateUnits.Months, 0, TCategoryAxisOptions.AutoMin | TCategoryAxisOptions.AutoMax | TCategoryAxisOptions.DateAxis | TCategoryAxisOptions.AutoCrossDate | TCategoryAxisOptions.AutoDate, AxisFont, "yyyy\\-mm\\-dd;@", true, AxisLine, AxisTicks, AxisRangeOptions, null, TChartAxisPos.Bottom, 1);
            AxisFont = new TFlxChartFont("Calibri", 180, TExcelColor.FromArgb(0x59, 0x59, 0x59), TFlxFontStyles.None, TFlxUnderline.None, TFontScheme.Minor);
            AxisLine = new TAxisLineOptions();
            AxisLine.MainAxis = new ChartLineOptions(new TShapeLine(true, new TLineStyle(new TSolidFill(TDrawingColor.FromTheme(TThemeColor.Dark1, new TColorTransform(TColorTransformType.LumMod, 0.15), new TColorTransform(TColorTransformType.LumOff, 0.85))), 9525, TPenAlignment.Center, TLineCap.Flat, TCompoundLineType.Single, null, TLineJoin.Round, null, null, null), null, TFormattingType.Subtle));
            AxisLine.MajorGridLines = new ChartLineOptions(new TShapeLine(true, new TLineStyle(new TSolidFill(TDrawingColor.FromTheme(TThemeColor.Dark1, new TColorTransform(TColorTransformType.LumMod, 0.15), new TColorTransform(TColorTransformType.LumOff, 0.85))), 9525, TPenAlignment.Center, TLineCap.Flat, TCompoundLineType.Single, null, TLineJoin.Round, null, null, null), null, TFormattingType.Subtle));
            AxisLine.DoNotDrawLabelsIfNotDrawingAxis = false;
            AxisTicks = new TAxisTickOptions(TTickType.None, TTickType.None, TAxisLabelPosition.NextToAxis, TBackgroundMode.Transparent, TDrawingColor.FromRgb(0x59, 0x59, 0x59), 0);
            CatAxis.NumberFormat = "yyyy-mm";
            CatAxis.NumberFormatLinkedToSource = false;

            TBaseAxis ValAxis = new TValueAxis(0, 0, 0, 0, 0, TValueAxisOptions.AutoMin | TValueAxisOptions.AutoMax | TValueAxisOptions.AutoMajor | TValueAxisOptions.AutoMinor | TValueAxisOptions.AutoCross, AxisFont, "General", true, AxisLine, AxisTicks, null, TChartAxisPos.Left);
            Chart1.SetChartAxis(new TChartAxis(0, CatAxis, ValAxis));

        }

        private void NormalOpen(ExcelFile Xls)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Xls.Save(saveFileDialog1.FileName);

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
}
