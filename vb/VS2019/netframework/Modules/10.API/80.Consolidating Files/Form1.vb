Imports System.Collections
Imports System.ComponentModel
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports System.IO
Imports System.Reflection

Namespace ConsolidatingFiles
	''' <summary>
	''' A demo on how to copy many sheets from different files into one file.
	''' </summary>
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Public Sub New()
			InitializeComponent()
		End Sub

		''' <summary>
		''' This is the method that will be called by the ASP.NET front end. It returns an array of bytes 
		''' with the report data, so the ASP.NET application can stream it to the client.
		''' </summary>
		''' <param name="fileDatas"></param>
		''' <param name="fileNames"></param>
		''' <param name="OnlyData"></param>
		''' <returns>The generated file as a byte array.</returns>
		Public Function WebRun(ByVal fileDatas() As Stream, ByVal fileNames() As String, ByVal OnlyData As Boolean) As Byte()
			If fileNames.Length <= 0 Then
				Throw New ApplicationException("You must select at least one file")
			End If

			Dim XlsOut As ExcelFile = Consolidate(fileDatas, fileNames, OnlyData)

			Using OutStream As New MemoryStream()
				XlsOut.Save(OutStream)
				Return OutStream.ToArray()
			End Using


		End Function

		Private Function Consolidate(ByVal fileDatas() As Stream, ByVal fileNames() As String, ByVal OnlyData As Boolean) As ExcelFile
			Dim XlsIn As ExcelFile = New XlsFile()
			Dim XlsOut As ExcelFile = New XlsFile(True)
			XlsOut.NewFile(1, TExcelFileFormat.v2019)

			If fileNames.Length > 1 AndAlso cbOnlyData.Checked Then
				XlsOut.InsertAndCopySheets(1, 2, fileNames.Length - 1)
			End If

			For i As Integer = 0 To fileNames.Length - 1
				If fileDatas IsNot Nothing Then
					XlsIn.Open(fileDatas(i))
				Else
					XlsIn.Open(fileNames(i))
				End If
				XlsIn.ConvertFormulasToValues(True) 'If there is any formula referring to other sheet, convert it to value.
													 'We could also call an overloaded version of InsertAndCopySheets() that
													 'copies many sheets at the same time, so references are kept.
				XlsOut.ActiveSheet = i + 1

				If OnlyData Then
					XlsOut.InsertAndCopyRange(TXlsCellRange.FullRange(), 1, 1, 1, TFlxInsertMode.ShiftRangeDown, TRangeCopyMode.All, XlsIn, 1)
				Else
					XlsOut.InsertAndCopySheets(1, XlsOut.ActiveSheet, 1, XlsIn)
				End If

				'Change sheet name.
				Dim s As String = Path.GetFileName(fileNames(i))
				If s.Length > 32 Then
					XlsOut.SheetName = s.Substring(0, 29) & "..."
				Else
					XlsOut.SheetName = s
				End If

			Next i

			If Not cbOnlyData.Checked Then
				XlsOut.ActiveSheet = XlsOut.SheetCount
				XlsOut.DeleteSheet(1) 'Remove the empty sheet that came with the workbook.
			End If

			XlsOut.ActiveSheet = 1
			Return XlsOut
		End Function

		Private Sub button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles button1.Click
			If openFileDialog1.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
				Return
			End If
			Dim fileNames() As String = openFileDialog1.FileNames
			If fileNames.Length <= 0 Then
				MessageBox.Show("You must select at least one file")
				Return
			End If

			Dim XlsOut As ExcelFile = Consolidate(Nothing, fileNames, cbOnlyData.Checked)

			If saveFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
				XlsOut.Save(saveFileDialog1.FileName)

				If MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
					Process.Start(saveFileDialog1.FileName)
				End If
			End If

		End Sub
	End Class
End Namespace
