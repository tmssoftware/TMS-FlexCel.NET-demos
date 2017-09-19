Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Report
Imports System.Transactions
Imports System.Configuration


Namespace EntityFramework
	''' <summary>
	''' Summary description for Form1.
	''' </summary>
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles button1.Click
			AutoRun()
		End Sub

		Public Sub AutoRun()
			Using ordersReport As New FlexCelReport(True)
				Dim DataPath As String = Path.Combine(Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), ".."), "..") & Path.DirectorySeparatorChar
				ordersReport.SetValue("Date", Date.Now)

				Using Northwind As New northwndEntities()
					ordersReport.AddTable("Categories", Northwind.Categories)
					ordersReport.AddTable("Products", Northwind.Products)

					If saveFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
						Dim transactionOptions As New TransactionOptions()
						transactionOptions.IsolationLevel = System.Transactions.IsolationLevel.Serializable 'it would be better to sue Snapshot here, but it isn't supported by SQL Sever CE
						Using transactionScope As New TransactionScope(TransactionScopeOption.Required, transactionOptions)
							ordersReport.Run(DataPath & "Entity Framework.template.xls", saveFileDialog1.FileName)
							transactionScope.Complete()
						End Using

						If MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
							Process.Start(saveFileDialog1.FileName)
						End If
					End If
				End Using
			End Using
		End Sub

		Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
			Close()
		End Sub
	End Class

End Namespace
