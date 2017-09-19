Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Report
Imports FlexCel.Demo.SharedData


Namespace PivotTables
	''' <summary>
	''' How to keep pivot tabels in reports.
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
				ordersReport.SetValue("Date", Date.Now)

				Dim DataPath As String = Path.Combine(Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), ".."), "..") & Path.DirectorySeparatorChar

				If saveFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
					Dim Template As String
					If Path.GetExtension(saveFileDialog1.FileName) = ".xlsx" Then Template = "Pivot Tables.template.xlsx" Else Template = "Pivot Tables.template.xls"
					ordersReport.Run(DataPath & Template, saveFileDialog1.FileName)

					If MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
						Process.Start(saveFileDialog1.FileName)
					End If
				End If
			End Using

		End Sub

		Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
			Close()
		End Sub
	End Class

End Namespace
