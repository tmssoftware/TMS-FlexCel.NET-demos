Imports System.Collections
Imports System.ComponentModel

Imports FlexCel.Core
Imports FlexCel.XlsAdapter

Namespace VirtualMode
	''' <summary>
	''' A demo on how to read a file from FlexCel and display the results.
	''' </summary>
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Private CellData As SparseCellArray 'we will store the data here. This is an example, in real world you would use "Virtual mode" to load the cells into your own structures.

		Public Sub New()
			InitializeComponent()
			ResizeToolbar(mainToolbar)
		End Sub

		Private Sub ResizeToolbar(ByVal toolbar As ToolStrip)

			Using gr As Graphics = CreateGraphics()
				Dim xFactor As Double = gr.DpiX / 96.0
				Dim yFactor As Double = gr.DpiY / 96.0
				toolbar.ImageScalingSize = New Size(CInt(Fix(24 * xFactor)), CInt(Fix(24 * yFactor)))
				toolbar.Width = 0 'force a recalc of the buttons.
			End Using
		End Sub

		Private Sub btnExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.Click
			Close()
		End Sub

		Private Sub btnInfo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnInfo.Click
			MessageBox.Show("This demo shows how to read the contents of an xls file without loading the file in memory." & vbLf & "We will first load the sheet names in the file, then open just a single sheet, and read all or just the 50 first rows of it.")
		End Sub


		Private Sub btnOpenFile_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOpenFile.Click
			If openFileDialog1.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
				Return
			End If
			ImportFile(openFileDialog1.FileName)
		End Sub

		Private Sub ImportFile(ByVal FileName As String)
			Try
				Dim xls As New XlsFile()
				xls.VirtualMode = True 'Remember to turn virtual mode on, or the event won't be called.

				'By default, FlexCel returns the formula text for the formulas, besides its calculated value.
				'If you are not interested in formula texts, you can gain a little performance by ignoring it.
				'This also works in non virtual mode.
				xls.IgnoreFormulaText = cbIgnoreFormulaText.Checked

				CellData = New SparseCellArray()

				'Attach the CellReader handler.
				Dim cr As New CellReader(cbFirst50Rows.Checked, CellData, cbFormatValues.Checked)
				AddHandler xls.VirtualCellStartReading, AddressOf cr.OnStartReading
				AddHandler xls.VirtualCellRead, AddressOf cr.OnCellRead

				Dim StartOpen As Date = Date.Now

				'Open the file. As we have a CellReader attached, the cells won't be loaded into memory, they will be passed to the CellReader
				xls.Open(FileName)
				Dim StartSheetSelect As Date = cr.StartSheetSelect
				Dim EndSheetSelect As Date = cr.EndSheetSelect

				Dim EndOpen As Date = Date.Now
				statusBar.Text = "Time to open file: " & (StartSheetSelect.Subtract(StartOpen)).ToString() & "     Time to load file and fill grid: " & (EndOpen.Subtract(EndSheetSelect)).ToString()

				'Set up grid.
				GridCaption.Text = FileName
				If CellData IsNot Nothing Then
					DisplayGrid.ColumnCount = CellData.ColCount
					DisplayGrid.RowCount = CellData.RowCount
				Else
					DisplayGrid.ColumnCount = 0
					DisplayGrid.RowCount = 0
				End If

				For i As Integer = 0 To DisplayGrid.ColumnCount - 1
					DisplayGrid.Columns(i).Name = TCellAddress.EncodeColumn(i + 1)
				Next i
			Catch
				GridCaption.Text = "Error Loading File"
				CellData = Nothing
				Throw
			End Try
		End Sub

		Private Sub DisplayGrid_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles DisplayGrid.RowPostPaint
			'Show the row number in the grid at the left
			Dim r As String = (e.RowIndex + 1).ToString()
			Dim textSize As SizeF = e.Graphics.MeasureString(r, DisplayGrid.Font)
			If DisplayGrid.RowHeadersWidth < CInt(textSize.Width + 20) Then
				DisplayGrid.RowHeadersWidth = CInt(textSize.Width + 20)
			End If
			e.Graphics.DrawString(r, DisplayGrid.Font, SystemBrushes.ControlText, e.RowBounds.Left + DisplayGrid.RowHeadersWidth - textSize.Width - 5, e.RowBounds.Location.Y + ((e.RowBounds.Height - textSize.Height) / 2F))
		End Sub

		Private Sub DisplayGrid_CellValueNeeded(ByVal sender As Object, ByVal e As DataGridViewCellValueEventArgs) Handles DisplayGrid.CellValueNeeded
			If CellData Is Nothing Then
				e.Value = Nothing
				Return
			End If

			If e.RowIndex >= CellData.RowCount Then
				e.Value = Nothing
				Return
			End If

			e.Value = CellData.GetValue(e.RowIndex + 1, e.ColumnIndex + 1)
		End Sub
	End Class
End Namespace
