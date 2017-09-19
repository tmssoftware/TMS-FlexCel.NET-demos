Imports System.Text
Imports FlexCel.Core

Namespace VirtualMode
	'A simple cell reader that will get the values from FlexCel and put them into a grid.
	Friend Class CellReader
		Private Only50Rows As Boolean
		Private CellData As SparseCellArray
		Private FormatValues As Boolean
		Private SheetToRead As Integer
		Public StartSheetSelect As Date
		Public EndSheetSelect As Date

		Public Sub New(ByVal aOnly50Rows As Boolean, ByVal aCellData As SparseCellArray, ByVal aFormatValues As Boolean)
			Only50Rows = aOnly50Rows
			CellData = aCellData
			FormatValues = aFormatValues
		End Sub

		Public Sub OnStartReading(ByVal sender As Object, ByVal e As VirtualCellStartReadingEventArgs)
			StartSheetSelect = Date.Now
			Using SheetSelector As New SheetSelectorForm(e.SheetNames)

				If Not SheetSelector.Execute() Then
					EndSheetSelect = Date.Now
					e.NextSheet = Nothing 'stop reading
					Return
				End If

				EndSheetSelect = Date.Now
				e.NextSheet = SheetSelector.SelectedSheet
				SheetToRead = SheetSelector.SelectedSheetIndex + 1
			End Using
		End Sub

		Public Sub OnCellRead(ByVal sender As Object, ByVal e As VirtualCellReadEventArgs)
			If Only50Rows AndAlso e.Cell.Row > 50 Then
				e.NextSheet = Nothing 'Stop reading all sheets.
				Return
			End If

			If e.Cell.Sheet <> SheetToRead Then
				e.NextSheet = Nothing 'Stop reading all sheets.
				Return
			End If

			If FormatValues Then
				Dim Clr As TUIColor = Color.Empty
				CellData.AddValue(e.Cell.Row, e.Cell.Col, TFlxNumberFormat.FormatValue(e.Cell.Value, CType(sender, ExcelFile).GetFormat(e.Cell.XF).Format, Clr, (CType(sender, ExcelFile))))
			Else
				CellData.AddValue(e.Cell.Row, e.Cell.Col, Convert.ToString(e.Cell.Value))
			End If
		End Sub
	End Class
End Namespace
