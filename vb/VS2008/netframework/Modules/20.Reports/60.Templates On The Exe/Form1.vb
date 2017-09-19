Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports System.Resources
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Report
Imports FlexCel.Demo.SharedData


Namespace TemplatesOnTheExe
	''' <summary>
	''' How to embed the reports in the executable, inlcluding "included" reports.
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
				AddHandler ordersReport.GetInclude, AddressOf ordersReport_GetInclude
				ordersReport.SetValue("Date", Date.Now)

				If saveFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
					Dim a As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
					Using InStream As Stream = a.GetManifestResourceStream("Templates On The Exe.template.xls")
						Using OutStream As New FileStream(saveFileDialog1.FileName, FileMode.Create)
							ordersReport.Run(InStream, OutStream)
						End Using
					End Using

					If MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
						Process.Start(saveFileDialog1.FileName)
					End If
				End If
			End Using

		End Sub

		Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
			Close()
		End Sub

		Private Sub ordersReport_GetInclude(ByVal sender As Object, ByVal e As FlexCel.Report.GetIncludeEventArgs)
			Dim a As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
            Using InStream As Stream = a.GetManifestResourceStream(e.FileName)
                Dim data(CInt(InStream.Length) - 1) As Byte
				InStream.Position = 0
				InStream.Read(data, 0, data.Length)
				e.IncludeData = data
			End Using
		End Sub
	End Class

End Namespace
