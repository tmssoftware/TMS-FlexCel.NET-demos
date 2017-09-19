Imports System.Collections
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports System.IO
Imports System.Reflection
Imports System.Text

Namespace AdvancedAPI
	''' <summary>
	''' A demo on creating a file using more advanced features.
	''' </summary>
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles button1.Click
			Dim Xls As ExcelFile = New XlsFile(True)
			AddData(Xls)

			NormalOpen(Xls)
		End Sub

		''' <summary>
		''' We will use this path to find the template.xls. Code is a little complex because it has to run in mono.
		''' </summary>
		Private ReadOnly Property PathToExe() As String
			Get
				Return Path.Combine(Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), ".."), "..") & Path.DirectorySeparatorChar
			End Get
		End Property

		'some silly data to fill in the cells. A real app would read this from somewhere else.
		Private Country() As String = { "USA", "Canada", "Spain", "France", "United Kingdom", "Australia", "Brazil", "Unknown" }

		Private DataRows As Integer = 100

		''' <summary>
		''' Will return a list of countries separated by Character(0) so it can be used as input for a built in list.
		''' </summary>
		''' <returns></returns>
		Private Function GetCountryList() As String
			Dim sb As New StringBuilder()
			Dim sep As String = ""
			For Each c As String In Country
				sb.Append(sep)
				sb.Append(c)
				sep = vbNullChar 'not very efficient method to concat, but good enough for this demo.
			Next c

			Return sb.ToString()
		End Function

		Private Sub AddChart(ByVal DataCell As TXlsNamedRange, ByVal Xls As ExcelFile)
			'Find the cell where the cart will go.
			Dim ChartRange As TXlsNamedRange = Xls.GetNamedRange("ChartData", -1)

			'Insert cells to expand the range for the chart. It already has 2 rows, so we need to insert Country.Length - 2
			'Note also that we insert after ChartRange.Top, so the chart is updates with the new range.
			Xls.InsertAndCopyRange(New TXlsCellRange(ChartRange.Top, ChartRange.Left, ChartRange.Top, ChartRange.Left + 1), ChartRange.Top + 1, ChartRange.Left, Country.Length - 2, TFlxInsertMode.ShiftRangeDown) 'we use shiftrangedown so not all the row goes down and the chart stays in place.

			'Get the cell addresses of the data range.
			Dim FirstCell As New TCellAddress(DataCell.Top, DataCell.Left)
			Dim SecondCell As New TCellAddress(DataCell.Top + DataRows, DataCell.Left + 1)
			Dim FirstSumCell As New TCellAddress(DataCell.Top, DataCell.Left + 1)

			'Fill a table with the data to be used in the chart.
			For r As Integer = ChartRange.Top To ChartRange.Top + Country.Length - 1
				Xls.SetCellValue(r, ChartRange.Left, Country(r - ChartRange.Top))
				Xls.SetCellValue(r, ChartRange.Left + 1, New TFormula("=SUMIF(" & FirstCell.CellRef & ":" & SecondCell.CellRef & ",""" & Country(r - ChartRange.Top) & """, " & FirstSumCell.CellRef & ":" & SecondCell.CellRef & ")"))
			Next r

		End Sub

		Private Sub AddData(ByVal Xls As ExcelFile)
			Dim TemplateFile As String = "template.xls"
			If cbXlsxTemplate.Checked Then
				If Not XlsFile.SupportsXlsx Then
					Throw New Exception("Xlsx files are not supported in this version of the .NET framework")
				End If
				TemplateFile = "template.xlsm"
			End If

			' Open an existing file to be used as template. In this example this file has
			' little data, in a real situation it should have as much as possible. (Or even better, be a report)
			Xls.Open(Path.Combine(PathToExe, TemplateFile))

			'Find the cell where we want to fill the data. In this case, we have created a named range "data" so the address
			'is not hardcoded here.
			Dim DataCell As TXlsNamedRange = Xls.GetNamedRange("Data", -1)

			'Add a chart with totals
			AddChart(DataCell, Xls)
			'Note that "DataCell" will change because we inserted rows above it when creating the chart. But we will keep using the old one.

			'Add the captions. This should probably go into the template, but in a dynamic environment it might go here.
			Xls.SetCellValue(DataCell.Top - 1, DataCell.Left, "Country")
			Xls.SetCellValue(DataCell.Top - 1, DataCell.Left + 1, "Quantity")

			'Add a rectangle around the cells
			Dim ApplyFormat As New TFlxApplyFormat()
			ApplyFormat.SetAllMembers(False)
			ApplyFormat.Borders.SetAllMembers(True) 'We will only apply the borders to the existing cell formats
			Dim fmt As TFlxFormat = Xls.GetDefaultFormat
			fmt.Borders.Left.Style = TFlxBorderStyle.Double
			fmt.Borders.Left.Color = Color.Black
			fmt.Borders.Right.Style = TFlxBorderStyle.Double
			fmt.Borders.Right.Color = Color.Black
			fmt.Borders.Top.Style = TFlxBorderStyle.Double
			fmt.Borders.Top.Color = Color.Black
			fmt.Borders.Bottom.Style = TFlxBorderStyle.Double
			fmt.Borders.Bottom.Color = Color.Black
			Xls.SetCellFormat(DataCell.Top - 1, DataCell.Left, DataCell.Top, DataCell.Left + 1, fmt, ApplyFormat, True) 'Set last parameter to true so it draws a box.

			'Freeze panes
			Xls.FreezePanes(New TCellAddress(DataCell.Top, 1))


			Dim Rnd As New Random()

			'Fill the data
			Dim z As Integer = 0
			Dim OutlineLevel As Integer = 0
			For r As Integer = 0 To DataRows

				'Fill the values.
				Xls.SetCellValue(DataCell.Top + r, DataCell.Left, Country(z Mod Country.Length)) 'For non C# users, "%" means "mod" or modulus in other languages. It is the rest of the integer division.
				Xls.SetCellValue(DataCell.Top + r, DataCell.Left + 1, Rnd.Next(1000))

				'Add the country to the outline
				Xls.SetRowOutlineLevel(DataCell.Top + r, OutlineLevel)
				'increment the country randomly
				If Rnd.Next(3) = 0 Then
					z += 1
					OutlineLevel = 0 'Break the group and create a new one.
				Else
					OutlineLevel = 1
				End If
			Next r

			'Make the "+" signs of the outline appear at the top.
			Xls.OutlineSummaryRowsBelowDetail = False

			'Collapse the outline to the first level.
			Xls.CollapseOutlineRows(1, TCollapseChildrenMode.Collapsed)

			'Add Data Validation for the first column, it must be a country.
			Dim dv As New TDataValidationInfo(TDataValidationDataType.List, TDataValidationConditionType.Between, "=""" & GetCountryList() & """", Nothing, False, True, True, True, "Unknown country", "Please make sure that the country is in the list", False, Nothing, Nothing, TDataValidationIcon.Stop) 'We will use the stop icon so no invalid input is permitted. - We will not use an input box, so this is false and the 2 next entries are null - Note that as we entered the data directly in FirstFormula, we need to set this to true - no need for a second formula, not used in List - We could have used a range of cells here with the values (like "=C1..C4") Instead, we directly entered the list in the formula. - This parameter does not matter since it is a list. It will not be used. - We will use a built in list.
			Xls.AddDataValidation(New TXlsCellRange(DataCell.Top, DataCell.Left, DataCell.Top + DataRows, DataCell.Left), dv)

			'Add Data Validation for the second column, it must be an integer between 0 and 1000.
			dv = New TDataValidationInfo(TDataValidationDataType.WholeNumber, TDataValidationConditionType.Between, "=0", "=1000", False, False, False, True, "Invalid Quantity", Nothing, True, "Quantity:", "Please enter a quantity between 0 and 1000", TDataValidationIcon.Stop) 'We will use the stop icon so no invalid input is permitted. - We will leave the default error message. - Second formula is the second part. - First formula marks the first part of the "between" condition. - We will request a number.
			Xls.AddDataValidation(New TXlsCellRange(DataCell.Top, DataCell.Left + 1, DataCell.Top + DataRows, DataCell.Left + 1), dv)


			'Search country "Unknown" and replace it by "no".
			'This does not make any sense here (we could just have entered "no" to begin)
			'but it shows how to do it when modifying an existing file
			Xls.Replace("Unknown", "no", TXlsCellRange.FullRange(), True, False, True)

			'Autofit the rows. As we keep the row height automatic this will not show when opening in Excel, but will work when directly printing from FlexCel.
			Xls.AutofitRowsOnWorkbook(False, True, 1)

			Xls.Recalc() 'Calculate the SUMIF formulas so we can sort by them. Note that FlexCel automatically recalculates before saving,
						  'but in this case we haven't saved yet, so the sheet is not recalculated. You do not normally need to call Recalc directly.

			'Sort the data. As in the case with replace, this does not make much sense. We could have entered the data sorted to begin
			'But it shows how you can use the feature.

			'Find the cell where the chart goes.
			Dim ChartRange As TXlsNamedRange = Xls.GetNamedRange("ChartData", -1)
			Xls.Sort(New TXlsCellRange(ChartRange.Top, ChartRange.Left, ChartRange.Top + Country.Length, ChartRange.Left + 1), True, New Integer() { 2 }, New TSortOrder() { TSortOrder.Descending }, Nothing)



			'Protect the Sheet
			Dim Sp As New TSheetProtectionOptions(False) 'Create default protection options that allows everything.
			Sp.InsertColumns = False 'Restrict inserting columns.
			Xls.Protection.SetSheetProtection("flexcel", Sp)
			'Set a modify password. Note that this does *not* encrypt the file.
			Xls.Protection.SetModifyPassword("flexcel", True, "flexcel")

			Xls.Protection.OpenPassword = "flexcel" 'OpenPasword is the only password that will actually encrypt the file, so you will not be able to open it with flexcel if you do not know the password.

			'Select cell A1
			Xls.SelectCell(1, 1, True)
		End Sub

		'This is part of an advanced feature (showing the user using a file) , you do not need to use
		'this method on normal places.
		Private Function GetLockingUser(ByVal FileName As String) As String
			Try
				Dim xerr As New XlsFile()
				xerr.Open(FileName)
				Return " - File might be in use by: " & xerr.Protection.WriteAccess
			Catch
				Return String.Empty
			End Try
		End Function

		Private Sub NormalOpen(ByVal Xls As ExcelFile)
			If cbXlsxTemplate.Checked Then
				saveFileDialog1.FilterIndex = 1
			Else
				saveFileDialog1.FilterIndex = 0
			End If
			If saveFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
				If (Not XlsFile.SupportsXlsx) AndAlso Path.GetExtension(saveFileDialog1.FileName) = ".xlsm" Then
					Throw New Exception("Xlsx files are not supported in this version of the .NET framework")
				End If


				Try
					Xls.Save(saveFileDialog1.FileName)
				Catch ex As IOException 'This is not really needed, just to show the username of the user locking the file.
					Throw New IOException(ex.Message & GetLockingUser(saveFileDialog1.FileName), ex)
				End Try

				If MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
					Process.Start(saveFileDialog1.FileName)
				End If
			End If
		End Sub
	End Class
End Namespace
