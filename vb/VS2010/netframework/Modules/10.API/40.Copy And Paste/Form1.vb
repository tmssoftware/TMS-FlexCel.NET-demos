Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports System.Text
Imports FlexCel.Core
Imports FlexCel.XlsAdapter

Namespace CopyAndPaste
	''' <summary>
	''' Copy and Paste Example.
	''' </summary>
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Xls As XlsFile

		Private Sub btnNewFile_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNewFile.Click
			Try
				Xls = New XlsFile()
				Xls.NewFile(1, TExcelFileFormat.v2016)
			Catch ex As Exception
				MessageBox.Show(ex.Message)
			End Try
		End Sub

		Private Sub btnOpenFile_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnOpenFile.Click
			Try
				If openFileDialog.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
					Return
				End If
				Xls = New XlsFile(openFileDialog.FileName)
			Catch ex As Exception
				MessageBox.Show(ex.Message)
			End Try

		End Sub


		Private Sub DoPaste(ByVal iData As IDataObject)
			If Xls Is Nothing Then
				MessageBox.Show("Please push the New File button before pasting")
				Return
			End If

			Try
				If iData.GetDataPresent(FlexCelDataFormats.Excel97) Then
					'DO NOT CALL -> using (MemoryStream ms = (MemoryStream)iData.GetData(FlexCelDataFormats.Excel97))
					'You shouldn't dispose the stream, as it belongs to the Clipboard.
					Dim o As Object = iData.GetData(FlexCelDataFormats.Excel97)
					Dim ms As MemoryStream = CType(o, MemoryStream)
						Xls.PasteFromXlsClipboardFormat(1, 1, TFlxInsertMode.NoneDown, ms)
						MessageBox.Show("NATIVE Data has been pasted at cell A1")
				Else
					If iData.GetDataPresent(DataFormats.UnicodeText) Then
					Xls.PasteFromTextClipboardFormat(1, 1, TFlxInsertMode.NoneDown, CStr(iData.GetData(DataFormats.UnicodeText)))
					MessageBox.Show("UNICODE TEXT Data has been pasted at cell A1")
				Else
						If iData.GetDataPresent(DataFormats.Text) Then
					Xls.PasteFromTextClipboardFormat(1, 1, TFlxInsertMode.NoneDown, CStr(iData.GetData(DataFormats.Text)))
					MessageBox.Show("TEXT Data has been pasted at cell A1")

				Else
					MessageBox.Show("There is no Excel or Text data on the clipboard")
				End If
				End If
				End If
			Catch ex As Exception
				MessageBox.Show(ex.Message)
				Xls = New XlsFile()
				Xls.NewFile()
			End Try
		End Sub

		Private Sub btnPaste_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPaste.Click
			DoPaste(Clipboard.GetDataObject())
		End Sub

		Private Sub DropHere_DragOver(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles DropHere.DragOver
			If e.Data.GetDataPresent(FlexCelDataFormats.Excel97) OrElse e.Data.GetDataPresent(DataFormats.UnicodeText) OrElse e.Data.GetDataPresent(DataFormats.Text) Then
				e.Effect = DragDropEffects.Copy
			End If
		End Sub


		Private Sub DropHere_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles DropHere.DragDrop
			DoPaste(e.Data)
		End Sub


		Private Sub DoCopy(ByVal ToClipboard As Boolean)
			If Xls Is Nothing Then
				MessageBox.Show("Please push the New File button before copying")
				Return
			End If

			'VERY IMPORTANT!!!!!
			'****************************************************************************
			'The MemoryStreams CAN NOT BE DISPOSED UNTIL WE CALL Clipboard.SetObjectData.
			'Even when we assigned the Stream with the DataObject Data, it is still in use and can't be freed.
			'****************************************************************************

			Try
				Dim data As New DataObject()
				Dim dataStreams As New List(Of MemoryStream)() 'we will use this list to dispose the memorystreams after they have been used.
				Try
					For Each cf As FlexCelClipboardFormat In System.Enum.GetValues(GetType(FlexCelClipboardFormat))
						Dim dataStream As New MemoryStream()
						dataStreams.Add(dataStream)
						Xls.CopyToClipboard(cf, dataStream)
						dataStream.Position = 0
						data.SetData(FlexCelDataFormats.GetString(cf), dataStream)

					Next cf
					If ToClipboard Then
						Clipboard.SetDataObject(data, True)
					Else
						DoDragDrop(data, DragDropEffects.Copy)
					End If
				Finally
					For Each ms As MemoryStream In dataStreams
						ms.Dispose()
					Next ms
				End Try
			Catch ex As Exception
				MessageBox.Show(ex.Message)
			End Try
		End Sub

		Private Sub btnCopy_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCopy.Click
			DoCopy(True)
		End Sub

		Private Sub btnDragMe_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles btnDragMe.MouseDown
			DoCopy(False)
		End Sub

	End Class
End Namespace
