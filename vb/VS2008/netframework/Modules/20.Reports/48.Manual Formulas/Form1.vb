Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Report


Namespace ManualFormulas
	''' <summary>
	''' Shows the formula tag.
	''' </summary>
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub SetupMines(ByVal MinesReport As FlexCelReport)
			Dim ds As New DataSet()
			Dim dtrows As DataTable = ds.Tables.Add("datarow")
			dtrows.Columns.Add("position", GetType(Integer))

			Dim dtcols As DataTable = ds.Tables.Add("datacol")
			dtcols.Columns.Add("position", GetType(Integer))
			dtcols.Columns.Add("value", GetType(Integer))

			ds.Relations.Add(dtrows.Columns("position"), dtcols.Columns("position"))

			'let's create 10 mines.
			Dim mines As New ArrayList()
			Dim rnd As New Random()
			Do While mines.Count < 10
				Dim nextMine As Integer = rnd.Next(9 * 9 - 1)
				Dim minepos As Integer = mines.BinarySearch(nextMine)
				If minepos >= 0 Then 'the value already exists
					Continue Do
				End If
				mines.Insert((Not minepos), nextMine)
			Loop

			'Fill the tables on master detail
			For r As Integer = 0 To 8
				dtrows.Rows.Add(New Object() { r })
				For c As Integer = 0 To 8
					Dim values(1) As Object
					values(0) = r
					If mines.BinarySearch(r * 9 + c) >= 0 Then
						values(1) = 1
					Else
						values(1) = DBNull.Value
					End If
					dtcols.Rows.Add(values)
				Next c
			Next r

			'finally, add the tables to the report.
			MinesReport.ClearTables()
			MinesReport.AddTable(ds, TDisposeMode.DisposeAfterRun) 'leave to Flexcel to delete the dataset.
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles button1.Click
			AutoRun()
		End Sub

		Public Sub AutoRun()
			Using MinesReport As New FlexCelReport(True)
				AddHandler MinesReport.AfterGenerateWorkbook, AddressOf MinesReport_AfterGenerateWorkbook
				SetupMines(MinesReport)
				Dim DataPath As String = Path.Combine(Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), ".."), "..") & Path.DirectorySeparatorChar

				If saveFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
					MinesReport.Run(DataPath & "Manual Formulas.template.xls", saveFileDialog1.FileName)

					If MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
						Process.Start(saveFileDialog1.FileName)
					End If
				End If
			End Using
		End Sub

		Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
			Close()
		End Sub

		Private Sub MinesReport_AfterGenerateWorkbook(ByVal sender As Object, ByVal e As FlexCel.Report.GenerateEventArgs)
			'do some "pretty" up for the final user.
			'we could do this directly on the template, but doing it here allows us to keep the template unprotected and easier to modify.

			e.File.ActiveSheet = 2
			e.File.SheetVisible = TXlsSheetVisible.Hidden
			e.File.ActiveSheet = 1
			e.File.Protection.SetSheetProtection(Nothing, New TSheetProtectionOptions(True))
			For r As Integer = 20 To FlxConsts.Max_Rows97_2003 + 1
				e.File.SetRowHidden(r, True)
			Next r
			For c As Integer = 12 To FlxConsts.Max_Columns97_2003 + 1
				e.File.SetColHidden(c, True)
			Next c
		End Sub
	End Class

End Namespace
