Imports System.Collections
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports System.IO
Imports System.Reflection
Imports System.Text

Namespace ChartAPI
	''' <summary>
	''' A demo on creating a chart with code.
	''' </summary>
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles button1.Click
			AutoRun()
		End Sub

		Private ReadOnly Property PathToExe() As String
			Get
				Return Path.Combine(Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), ".."), "..") & Path.DirectorySeparatorChar
			End Get
		End Property

		Private Sub AutoRun()
			'We will use data already stored in a file. For a real case, you would
			'probably fill this data from some database.
			Dim fileName As String = Path.Combine(PathToExe, "git-stats.xlsx")
			Dim Xls As ExcelFile = New XlsFile(fileName, True)

			'Add a new empty sheet for adding the chart.
			Xls.InsertAndCopySheets(0, 1, 1)
			Xls.ActiveSheet = 1
			Xls.SheetName = "Chart"
			Xls.PrintToFit = True
			Xls.PrintScale = 70
			Xls.PrintXResolution = 600
			Xls.PrintYResolution = 600
			Xls.PrintOptions = TPrintOptions.None
			Xls.PrintPaperSize = TPaperSize.Letter
			Xls.PrintLandscape = True

			AddChart(Xls)
			NormalOpen(Xls)
		End Sub



		Private Sub AddChart(ByVal xls As ExcelFile)
			'This code is adapted from APIMate.
			'Objects
			Dim ChartOptions1 As New TShapeProperties()
			ChartOptions1.Anchor = New TClientAnchor(TFlxAnchorType.MoveAndResize, 1, 215, 1, 608, 30, 228, 17, 736)
			ChartOptions1.ShapeName = "Lines of code"
			ChartOptions1.Print = True
			ChartOptions1.Visible = True
			ChartOptions1.ShapeOptions.SetValue(TShapeOption.fLockText, True)
			ChartOptions1.ShapeOptions.SetValue(TShapeOption.LockRotation, True)
			ChartOptions1.ShapeOptions.SetValue(TShapeOption.fAutoTextMargin, True)
			ChartOptions1.ShapeOptions.SetValue(TShapeOption.fillColor, 134217806)
			ChartOptions1.ShapeOptions.SetValue(TShapeOption.wzName, "Lines of code")
			Dim Chart1 As ExcelChart = xls.AddChart(ChartOptions1, TChartType.Area, New ChartStyle(102), False)

			Dim Title As New TDataLabel()
			Title.PositionZeroBased = Nothing
			Dim TextFillOptions As New ChartFillOptions(New TShapeFill(New TSolidFill(TDrawingColor.FromRgb(&H80, &H80, &H80)), True, TFormattingType.Subtle, TDrawingColor.FromRgb(&H0, &H0, &H0, New TColorTransform(TColorTransformType.Alpha, 0)), False))
			Dim LabelTextOptions As New TChartTextOptions(New TFlxChartFont("Calibri Light", 320, TExcelColor.FromArgb(&H80, &H80, &H80), TFlxFontStyles.Bold, TFlxUnderline.None, TFontScheme.Major), THFlxAlignment.center, TVFlxAlignment.center, TBackgroundMode.Transparent, TextFillOptions)
			Title.TextOptions = LabelTextOptions
			Dim LabelOptions As New TDataLabelOptions()
			Title.LabelOptions = LabelOptions
			Dim ChartLineOptions As New ChartLineOptions(New TShapeLine(True, New TLineStyle(New TNoFill(), Nothing), Nothing, TFormattingType.Subtle))
			Dim ChartFillOptions As New ChartFillOptions(New TShapeFill(New TNoFill(), False, TFormattingType.Subtle, Nothing, False))
			Title.Frame = New TChartFrameOptions(ChartLineOptions, ChartFillOptions, False)

			Dim Runs() As TRTFRun
			Runs = New TRTFRun(0){}
			Runs(0).FirstChar = 0
			Dim fnt As TFlxFont
			fnt = xls.GetDefaultFont
			fnt.Name = "Calibri Light"
			fnt.Size20 = 320
			fnt.Color = TExcelColor.FromArgb(&H80, &H80, &H80)
			fnt.Style = TFlxFontStyles.Bold
			fnt.Family = 0
			fnt.CharSet = 1
			fnt.Scheme = TFontScheme.Major
			Runs(0).FontIndex = xls.AddFont(fnt)
			Dim LabelValue1 As New TRichString("FlexCel: Lines of code over time", Runs, xls)

			Title.LabelValues = New Object() { LabelValue1 }

			Chart1.SetTitle(Title)

			Chart1.Background = New TChartFrameOptions(TDrawingColor.FromTheme(TThemeColor.Dark1, New TColorTransform(TColorTransformType.LumMod, 0.15), New TColorTransform(TColorTransformType.LumOff, 0.85)), 9525, TDrawingColor.FromTheme(TThemeColor.Light1), False)

			Dim PlotAreaFrame As TChartFrameOptions
			ChartLineOptions = New ChartLineOptions(New TShapeLine(True, New TLineStyle(New TNoFill(), Nothing), Nothing, TFormattingType.Subtle))
			ChartFillOptions = New ChartFillOptions(New TShapeFill(New TPatternFill(TDrawingColor.FromTheme(TThemeColor.Dark1, New TColorTransform(TColorTransformType.LumMod, 0.15), New TColorTransform(TColorTransformType.LumOff, 0.85)), TDrawingColor.FromTheme(TThemeColor.Light1), TDrawingPattern.ltDnDiag), True, TFormattingType.Subtle, Nothing, False))
			PlotAreaFrame = New TChartFrameOptions(ChartLineOptions, ChartFillOptions, False)
			Dim PlotAreaPos As New TChartPlotAreaPosition(True, TChartRelativeRectangle.Automatic, TChartLayoutTarget.Inner, True)
			Chart1.PlotArea = New TChartPlotArea(PlotAreaFrame, PlotAreaPos, False)

			Chart1.SetChartOptions(1, New TAreaChartOptions(False, TStackedMode.Stacked, Nothing))

			Dim LastYear As Integer = 0
			Dim shade As Double = 1
			For i As Integer = 2 To 189
				Dim Series As New ChartSeries("=" & (New TCellAddress("Data", 1, i, True, True)).CellRef, "=" & (New TCellAddress("Data", 2, i, True, True)).CellRef & ":" & (New TCellAddress("Data", 189, i, True, True)).CellRef, "=Data!$A$2:$A$189")

				'We will display every year in a single color. Each month gets its own shade.
				Dim xf As Integer = -1
				Dim Year As Integer = FlxDateTime.FromOADate((CDbl(xls.GetCellValue(2, 1, i, xf))), False).Year
				If LastYear <> Year Then
					shade = 1
				ElseIf shade > 0.3 Then
					shade -= 0.05
				End If
					LastYear = Year
				Dim SeriesColor As TDrawingColor = TDrawingColor.FromTheme(CType(TThemeColor.Accent1 + (Year Mod 6), TThemeColor), New TColorTransform(TColorTransformType.Shade, shade))


				Dim SeriesFill As New ChartSeriesFillOptions(New TShapeFill(New TSolidFill(SeriesColor), True, TFormattingType.Subtle, Nothing, False), Nothing, False, False)
				Dim SeriesLine As New ChartSeriesLineOptions(New TShapeLine(True, New TLineStyle(New TNoFill(), Nothing), Nothing, TFormattingType.Subtle), False)
				Series.Options.Add(New ChartSeriesOptions(-1, SeriesFill, SeriesLine, Nothing, Nothing, Nothing, True))

				Chart1.AddSeries(Series)
			Next i

			Chart1.PlotEmptyCells = TPlotEmptyCells.Zero
			Chart1.ShowDataInHiddenRowsAndCols = False

			Dim AxisFont As New TFlxChartFont("Calibri", 180, TExcelColor.FromArgb(&H59, &H59, &H59), TFlxFontStyles.None, TFlxUnderline.None, TFontScheme.Minor)
			Dim AxisLine As New TAxisLineOptions()
			AxisLine.MainAxis = New ChartLineOptions(New TShapeLine(True, New TLineStyle(New TSolidFill(TDrawingColor.FromTheme(TThemeColor.Dark1, New TColorTransform(TColorTransformType.LumMod, 0.15), New TColorTransform(TColorTransformType.LumOff, 0.85))), 9525, TPenAlignment.Center, TLineCap.Flat, TCompoundLineType.Single, Nothing, TLineJoin.Round, Nothing, Nothing, Nothing), Nothing, TFormattingType.Subtle))
			AxisLine.DoNotDrawLabelsIfNotDrawingAxis = False
			Dim AxisTicks As New TAxisTickOptions(TTickType.Outside, TTickType.None, TAxisLabelPosition.NextToAxis, TBackgroundMode.Transparent, TDrawingColor.FromRgb(&H59, &H59, &H59), 0)
			Dim AxisRangeOptions As New TAxisRangeOptions(12, 1, False, False, False)
			Dim CatAxis As TBaseAxis = New TCategoryAxis(0, 0, 12, TDateUnits.Days, 12, TDateUnits.Days, TDateUnits.Months, 0, TCategoryAxisOptions.AutoMin Or TCategoryAxisOptions.AutoMax Or TCategoryAxisOptions.DateAxis Or TCategoryAxisOptions.AutoCrossDate Or TCategoryAxisOptions.AutoDate, AxisFont, "yyyy\-mm\-dd;@", True, AxisLine, AxisTicks, AxisRangeOptions, Nothing, TChartAxisPos.Bottom, 1)
			AxisFont = New TFlxChartFont("Calibri", 180, TExcelColor.FromArgb(&H59, &H59, &H59), TFlxFontStyles.None, TFlxUnderline.None, TFontScheme.Minor)
			AxisLine = New TAxisLineOptions()
			AxisLine.MainAxis = New ChartLineOptions(New TShapeLine(True, New TLineStyle(New TSolidFill(TDrawingColor.FromTheme(TThemeColor.Dark1, New TColorTransform(TColorTransformType.LumMod, 0.15), New TColorTransform(TColorTransformType.LumOff, 0.85))), 9525, TPenAlignment.Center, TLineCap.Flat, TCompoundLineType.Single, Nothing, TLineJoin.Round, Nothing, Nothing, Nothing), Nothing, TFormattingType.Subtle))
			AxisLine.MajorGridLines = New ChartLineOptions(New TShapeLine(True, New TLineStyle(New TSolidFill(TDrawingColor.FromTheme(TThemeColor.Dark1, New TColorTransform(TColorTransformType.LumMod, 0.15), New TColorTransform(TColorTransformType.LumOff, 0.85))), 9525, TPenAlignment.Center, TLineCap.Flat, TCompoundLineType.Single, Nothing, TLineJoin.Round, Nothing, Nothing, Nothing), Nothing, TFormattingType.Subtle))
			AxisLine.DoNotDrawLabelsIfNotDrawingAxis = False
			AxisTicks = New TAxisTickOptions(TTickType.None, TTickType.None, TAxisLabelPosition.NextToAxis, TBackgroundMode.Transparent, TDrawingColor.FromRgb(&H59, &H59, &H59), 0)
			CatAxis.NumberFormat = "yyyy-mm"
			CatAxis.NumberFormatLinkedToSource = False

			Dim ValAxis As TBaseAxis = New TValueAxis(0, 0, 0, 0, 0, TValueAxisOptions.AutoMin Or TValueAxisOptions.AutoMax Or TValueAxisOptions.AutoMajor Or TValueAxisOptions.AutoMinor Or TValueAxisOptions.AutoCross, AxisFont, "General", True, AxisLine, AxisTicks, Nothing, TChartAxisPos.Left)
			Chart1.SetChartAxis(New TChartAxis(0, CatAxis, ValAxis))

		End Sub

		Private Sub NormalOpen(ByVal Xls As ExcelFile)
			If saveFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
				Xls.Save(saveFileDialog1.FileName)

				If MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
					Process.Start(saveFileDialog1.FileName)
				End If
			End If
		End Sub
	End Class
End Namespace
