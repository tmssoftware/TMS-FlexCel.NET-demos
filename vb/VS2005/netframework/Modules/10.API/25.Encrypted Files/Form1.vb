Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports System.Reflection

Imports FlexCel.Core
Imports FlexCel.XlsAdapter

Namespace EncryptedFiles
	''' <summary>
	''' Shows how to deal with Encrypted files.
	''' </summary>
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.Click
			Close()
		End Sub

		'The event that will actually provide the password to open the empty form.
		Private Sub GetPassword(ByVal e As OnPasswordEventArgs)
			Dim Pwd As New PasswordDialog()
			e.Password = String.Empty
			If Pwd.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
				Return
			End If
			e.Password = Pwd.Password
		End Sub

		Private Sub btnGo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnGo.Click
			' On this demo we will fill data on an existing file with the api, starting with an encrypted file holding the starting formats.

			' Declare some data for the chart.
			Dim Names() As String = { "Dog", "Cat", "Cow", "Horse", "Fish" }
			Dim Quantities() As Integer = { 123, 200, 150, 0, 180 }

			' Use two folders up to where the exe is to store the data. (Exe is stored at bin\debug)
			Dim DataPath As String = Path.Combine(Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), ".."), "..") & Path.DirectorySeparatorChar

			Dim xls As New XlsFile(True)

			' We will use the OnPassword event here to show how to 
			' open a file if you don't know a priory if it is encrypted or not.
			' If you already knew the file was encrypted, (as in this case)you could use:
			' xls.Protection.OpenPassword = "42";

			AddHandler xls.Protection.OnPassword, AddressOf GetPassword
			xls.Open(Path.Combine(DataPath, "EmptyForm.xls"))

			' Insert rows so the chart range grows. On this case we assume the data is at least 2 rows long. If not, we should handle 
			' the case and do a xls.DeleteRange.
			xls.InsertAndCopyRange(New TXlsCellRange(1, 1, 1, 2), 5, 1, Names.Length - 2, TFlxInsertMode.ShiftRangeDown, TRangeCopyMode.None)

			' Fill the data.
			For i As Integer = 0 To Names.Length - 1
				xls.SetCellValue(4 + i, 1, Names(i))
				xls.SetCellValue(4 + i, 2, Quantities(i))
			Next i

			' Set a new password for opening.
			xls.Protection.OpenPassword = "43"
			xls.Protection.SetModifyPassword("43", False, "Ford Prefect")

			If saveFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
				xls.Save(saveFileDialog1.FileName)

				If MessageBox.Show("Do you want to open the generated file? (Remember password is 43)", "Confirm", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
					Process.Start(saveFileDialog1.FileName)
				End If
			End If
		End Sub

	End Class
End Namespace
