Imports System.Collections
Imports System.ComponentModel

Imports FlexCel.Core
Imports FlexCel.XlsAdapter

Namespace ReadingFiles
	''' <summary>
	''' A demo on how to read a file from FlexCel and display the results.
	''' </summary>
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

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

		Private Sub btnOpenFile_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOpenFile.Click
			If openFileDialog1.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
				Return
			End If
			ImportFile(openFileDialog1.FileName, btnFormatValues.Checked)
		End Sub

		Private Sub ImportFile(ByVal FileName As String, ByVal Formatted As Boolean)
			Try
				'Open the Excel file.
				Dim xls As New XlsFile(False)
				Dim StartOpen As Date = Date.Now
				xls.Open(FileName)
				Dim EndOpen As Date = Date.Now

				'Set up the Grid
				DisplayGrid.DataBindings.Clear()
				DisplayGrid.DataSource = Nothing
				DisplayGrid.DataMember = Nothing
				Dim dataSet1 As New DataSet()
				sheetCombo.Items.Clear()

				'We will create a DataTable "SheetN" for each sheet on the Excel sheet.
				For sheet As Integer = 1 To xls.SheetCount
					xls.ActiveSheet = sheet

					sheetCombo.Items.Add(xls.SheetName)

					Dim Data As DataTable = dataSet1.Tables.Add("Sheet" & sheet.ToString())
					Data.BeginLoadData()
					Try
						Dim ColCount As Integer = xls.ColCount
						'Add one column on the dataset for each used column on Excel.
						For c As Integer = 1 To ColCount
							Data.Columns.Add(TCellAddress.EncodeColumn(c), GetType(String)) 'Here we will add all strings, since we do not know what we are waiting for.
						Next c

						Dim dr(ColCount - 1) As String

						Dim RowCount As Integer = xls.RowCount
						For r As Integer = 1 To RowCount
							Array.Clear(dr, 0, dr.Length)
							'This loop will only loop on used cells. It is more efficient than looping on all the columns.
							For cIndex As Integer = xls.ColCountInRow(r) To 1 Step -1 'reverse the loop to avoid calling ColCountInRow more than once.
								Dim Col As Integer = xls.ColFromIndex(r, cIndex)

								If Formatted Then
									Dim rs As TRichString = xls.GetStringFromCell(r, Col)
									dr(Col - 1) = rs.Value
								Else
									Dim XF As Integer = 0 'This is the cell format, we will not use it here.
									Dim val As Object = xls.GetCellValueIndexed(r, cIndex, XF)

									Dim Fmla As TFormula = TryCast(val, TFormula)
									If Fmla IsNot Nothing Then
										'When we have formulas, we want to write the formula result. 
										'If we wanted the formula text, we would not need this part.
										dr(Col - 1) = Convert.ToString(Fmla.Result)
									Else
										dr(Col - 1) = Convert.ToString(val)
									End If
								End If
							Next cIndex
							Data.Rows.Add(dr)
						Next r
					Finally
						Data.EndLoadData()
					End Try

					Dim EndFill As Date = Date.Now
					statusBar.Text = String.Format("Time to load file: {0}    Time to fill dataset: {1}     Total time: {2}", (EndOpen.Subtract(StartOpen)).ToString(), (EndFill.Subtract(EndOpen)).ToString(), (EndFill.Subtract(StartOpen)).ToString())

				Next sheet

				'Set up grid.
				DisplayGrid.DataSource = dataSet1
				DisplayGrid.DataMember = "Sheet1"
				sheetCombo.SelectedIndex = 0
				DisplayGrid.CaptionText = FileName

			Catch
				DisplayGrid.CaptionText = "Error Loading File"
				DisplayGrid.DataSource = Nothing
				DisplayGrid.DataMember = ""
				sheetCombo.Items.Clear()
				Throw
			End Try
		End Sub

		Private Sub sheetCombo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles sheetCombo.SelectedIndexChanged
			If (TryCast(sender, ComboBox)).SelectedIndex < 0 Then
				Return
			End If
			DisplayGrid.DataMember = "Sheet" & ((TryCast(sender, ComboBox)).SelectedIndex + 1).ToString()
		End Sub

		Private Sub AnalizeFile(ByVal FileName As String, ByVal Row As Integer, ByVal Col As Integer)
			Dim xls As New XlsFile()
			xls.Open(FileName)

			Dim XF As Integer = 0
			MessageBox.Show("Active sheet is """ & xls.ActiveSheetByName & """")
			Dim v As Object = xls.GetCellValue(Row, Col, XF)

			If v Is Nothing Then
				MessageBox.Show("Cell A1 is empty")
				Return
			End If

			'Here we have all the kind of objects FlexCel can return.
			Select Case Type.GetTypeCode(v.GetType())
				Case TypeCode.Boolean
					MessageBox.Show("Cell A1 is a boolean: " & CBool(v))
					Return
				Case TypeCode.Double 'Remember, dates are doubles with date format.
					Dim CellColor As TUIColor = Color.Empty
					Dim HasDate, HasTime As Boolean
					Dim CellValue As String = TFlxNumberFormat.FormatValue(v, xls.GetFormat(XF).Format, CellColor, xls, HasDate, HasTime).ToString()

					If HasDate OrElse HasTime Then
						MessageBox.Show("Cell A1 is a DateTime value: " & FlxDateTime.FromOADate(CDbl(v), xls.OptionsDates1904).ToString() & vbLf & "The value is displayed as: " & CellValue)
					Else
						MessageBox.Show("Cell A1 is a double: " & CDbl(v) & vbLf & "The value is displayed as: " & CellValue & vbLf)
					End If
					Return
				Case TypeCode.String
					MessageBox.Show("Cell A1 is a string: " & v.ToString())
					Return
			End Select

			Dim Fmla As TFormula = TryCast(v, TFormula)
			If Fmla IsNot Nothing Then
				MessageBox.Show("Cell A1 is a formula: " & Fmla.Text & "   Value: " & Convert.ToString(Fmla.Result))
				Return
			End If

			Dim RSt As TRichString = TryCast(v, TRichString)
			If RSt IsNot Nothing Then
				MessageBox.Show("Cell A1 is a formatted string: " & RSt.Value)
				Return
			End If

			If TypeOf v Is TFlxFormulaErrorValue Then
				MessageBox.Show("Cell A1 is an error: " & TFormulaMessages.ErrString(CType(v, TFlxFormulaErrorValue)))
				Return
			End If

			Throw New Exception("Unexpected value on cell")

		End Sub

		Private Sub btnInfo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnInfo.Click
			MessageBox.Show("This demo shows how to read the contents of an xls file" & vbLf & "The 'Open File' button will load an Excel file into a dataset. Depending on the button 'Format Values' it will load the actual values (this is the fastest) or the formatted values." & vbLf & "The 'Format Values' button will modify how the files are read when you press 'Open File'. Formated values are slower, but they will look just how Excel shows them." & vbLf & "The 'Value in Cell A1' button will load an Excel file and show the contents of cell a1 on the active sheet.")
		End Sub

		''' <summary>
		''' This method will not do anything truly useful, but it alows you to see how to 
		''' process the different types of objects that GetCellValue can return
		''' </summary>
		''' <param name="sender"></param>
		''' <param name="e"></param>
		Private Sub btnValueInCurrentCell_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnValueInCellA1.Click
			If openFileDialog1.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
				Return
			End If
			AnalizeFile(openFileDialog1.FileName, 1, 1)
		End Sub
	End Class
End Namespace
