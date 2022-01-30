Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Report
Imports FlexCel.Demo.SharedData


Namespace ChartsWithDynamicSeries
	''' <summary>
	''' A report including charts which have a series per row.
	''' </summary>
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles button1.Click
			AutoRun()
		End Sub

		Public Sub AutoRun()
			Using ordersReport As FlexCelReport = SharedData.CreateReport()
				AddHandler ordersReport.CustomizeChart, AddressOf OrdersReport_CustomizeChart
				Dim DataPath As String = Path.Combine(Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), ".."), "..") & Path.DirectorySeparatorChar

				If saveFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
					ordersReport.Run(DataPath & "Charts With Dynamic Series.template.xlsx", saveFileDialog1.FileName)

					If MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
						Process.Start(saveFileDialog1.FileName)
					End If
				End If
			End Using
		End Sub

		Private Sub OrdersReport_CustomizeChart(ByVal sender As Object, ByVal e As CustomizeChartEventArgs)
			If e.ChartName = "Stock<#swap series>" Then
				'In this event we will set the colors of the series depending on the product.
				'Let's image each product has an associated color that we want to use for its series.
				For subChart As Integer = 1 To e.Chart.SubchartCount
					For series As Integer = 1 To e.Chart.SeriesInSubchart(subChart)
						Dim seriesDef = e.Chart.GetSeriesInSubchart(subChart, series, True, True, True)
						Dim seriesOptions = seriesDef.Options(-1)
						seriesOptions.FillOptions = New ChartSeriesFillOptions(New TShapeFill(True, New TSolidFill(ColorForProduct(series))), Nothing, False, False)
						e.Chart.SetSeriesInSubchart(subChart, series, seriesDef)
					Next series
				Next subChart
			End If
		End Sub

		Private Shared Function ColorForProduct(ByVal series As Integer) As TDrawingColor
			Return TDrawingColor.FromRgb(CByte((series * 24) Mod 255), CByte((series * 32) Mod 255), CByte((series * 16) Mod 255))
		End Function

		Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
			Close()
		End Sub
	End Class

End Namespace
