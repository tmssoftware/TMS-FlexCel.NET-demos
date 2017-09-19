Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Report


Namespace DebuggingReports
	''' <summary>
	''' How to debug a report.
	''' </summary>
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Public Sub New()
			InitializeComponent()
		End Sub


		Private Sub button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles button1.Click
			AutoRun()
		End Sub

		Private Function CreateReport() As FlexCelReport
			Dim Result As New FlexCelReport(True)

			Result.SetValue("test", 3)
			Result.SetValue("tagval", 1)
			Result.SetValue("refval", "l")

			'Here we will add a dummy table with some fantasy values
			Dim dt As New DataTable("testdb")
			dt.Columns.Add("key", GetType(Integer))
			dt.Columns.Add("data", GetType(String))
			dt.Rows.Add(New Object() { 5, "cat" })
			dt.Rows.Add(New Object() { 6, "dog" })
			Result.AddTable("testdb", dt, TDisposeMode.DisposeAfterRun)

			Return Result
		End Function

		Public Sub AutoRun()
			Using Report As FlexCelReport = CreateReport()
				Dim DataPath As String = Path.Combine(Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), ".."), "..") & Path.DirectorySeparatorChar

				If saveFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
					Report.Run(DataPath & "Debugging Reports.template.xls", saveFileDialog1.FileName)

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
