Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports System.Resources
Imports System.Globalization
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Report
Imports FlexCel.Demo.SharedData


Namespace DirectSQL
	''' <summary>
	''' Summary description for Form1.
	''' </summary>
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles button1.Click
			Using genericReport As New FlexCelReport(True)
				Dim genericAdapter As IDbDataAdapter = SharedData.GetDataAdapter()
				Try
					genericReport.SetValue("ReportCaption", "Sales by Country and Employee")
					genericReport.AddConnection("Northwind", genericAdapter, CultureInfo.CurrentCulture)

					'In OleDb the parameters are positional, you don't really need to name them when creating them.
					'But when you are using an SQL Server connection, you *need*
					'to specify the parameter name ("@StartDate") and make it equal to "@" + the name
					'of the parameter. It is recommended that you always specify the name, even in OleDb connections.

					'Also, we are not going to create the parameters directly here (using new SqlCeParameter(...).
					'We are going to centralize all data access for the demos in SharedData, so we can change it and change all demos.
					genericReport.AddSqlParameter("StartDate", SharedData.CreateParameter("@StartDate", startDate.Value.Date))
					genericReport.AddSqlParameter("EndDate", SharedData.CreateParameter("@EndDate", endDate.Value.Date))
					Dim DataPath As String = Path.Combine(Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), ".."), "..") & Path.DirectorySeparatorChar

					If saveFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
						genericReport.Run(DataPath & "Direct SQL.template.xls", saveFileDialog1.FileName)

						If MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
							Process.Start(saveFileDialog1.FileName)
						End If
					End If
				Finally
					CType(genericAdapter, IDisposable).Dispose()
				End Try
			End Using
		End Sub

		Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
			Close()
		End Sub
	End Class

End Namespace
