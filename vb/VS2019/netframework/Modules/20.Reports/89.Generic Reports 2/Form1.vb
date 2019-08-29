Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports System.Data.OleDb
Imports System.Threading
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Report


Namespace GenericReports2
	''' <summary>
	''' A generic report.
	''' </summary>
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Private SqlDialog As EnterSQLDialog

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

		Private Sub button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.Click
			Close()
		End Sub

		Private Sub btnOpenconnection_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOpenConnection.Click
			Dim DataPath As String = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) & Path.DirectorySeparatorChar
			Dim ConfigFile As String = DataPath & "GenericReports2.udl"
			If Not File.Exists(ConfigFile) Then
				Using f As FileStream = File.Create(ConfigFile)
					'Nothing, create an empty udl.
				End Using
			End If

			Process.Start(ConfigFile)
		End Sub

		Private Sub btnQuery_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnQuery.Click
			Dim DataPath As String = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) & Path.DirectorySeparatorChar
			Dim ConfigFile As String = DataPath & "GenericReports2.udl"
			Connection.Close()
			dataSet = New DataSet()


			Connection.ConnectionString = "File Name = " & ConfigFile

			Connection.Open()

			If SqlDialog Is Nothing Then
				SqlDialog = New EnterSQLDialog()
			End If

			If SqlDialog.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
				Return
			End If

			dbDataAdapter.SelectCommand = New OleDbCommand(SqlDialog.SQL, Connection)
			dbDataAdapter.Fill(dataSet, "Table")
			dataGrid.CaptionText = dbDataAdapter.SelectCommand.CommandText
			dataGrid.SetDataBinding(dataSet, "Table")
		End Sub

		Private Sub Export(ByVal SQL As String, ByRef DataPath As String)
			Report.ClearTables()
			Report.AddTable(dataSet)
			Report.SetValue("Date", Date.Now)
			Report.SetValue("ReportCaption", SQL)

			DataPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) & Path.DirectorySeparatorChar 'First try to find the template on exe folder.

			If Not File.Exists(DataPath & "Generic Reports 2.template.xls") Then 'When on design mode, search for the template 2 folders up.
				DataPath = Path.Combine(DataPath, Path.Combine("..", "..")) & Path.DirectorySeparatorChar
			End If
		End Sub

		Private Sub btnExportExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExportExcel.Click
			Dim DataPath As String = Nothing
			If dbDataAdapter Is Nothing OrElse dbDataAdapter.SelectCommand Is Nothing OrElse dbDataAdapter.SelectCommand.CommandText Is Nothing Then
				MessageBox.Show("You need to select a query first")
				Return
			End If
			Export(dbDataAdapter.SelectCommand.CommandText, DataPath)

			If saveFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
				Report.Run(DataPath & "Generic Reports 2.template.xls", saveFileDialog1.FileName)

				If MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
					Process.Start(saveFileDialog1.FileName)
				End If
			End If
		End Sub
	End Class
End Namespace
