Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Report


Namespace Overflowsheets
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles button1.Click
			AutoRun()
		End Sub

		Public Sub AutoRun()
			Using Report As New FlexCelReport(True)
				Dim Data(1009) As TMyData
				For i As Integer = 0 To Data.Length - 1
					Data(i) = New TMyData("Customer " & i.ToString())
				Next i
				Report.AddTable("data", Data)
				Report.SetValue("split", 40)

				Dim DataPath As String = Path.Combine(Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), ".."), "..") & Path.DirectorySeparatorChar

				If saveFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
					Report.Run(DataPath & "Overflow Sheets.template.xlsx", saveFileDialog1.FileName)

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

	Friend Class TMyData
		Public Property Name() As String

		Public Sub New(ByVal name As String)
			Me.Name = name
		End Sub
	End Class

End Namespace
