Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports System.Resources
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Report
Imports FlexCel.Demo.SharedData


Namespace MasterDetailononeTable
	''' <summary>
	''' How to split a table into 2.
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
			Using ordersReport As FlexCelReport = SharedData.CreateReport()
				ordersReport.SetValue("Date", Date.Now)
				ordersReport.SetValue("ReportCaption", "Sales by year and country")

				Using ds As New DataSet()
					SharedData.Fill(ds, "SELECT Employees.Country, SUM([Order Details].UnitPrice * [Order Details].Quantity) AS Sales, COUNT([Order Details].Quantity) AS OrderCount, DatePart(yyyy, Orders.OrderDate) AS SaleYear, DatePart(q, Orders.OrderDate) AS Quarter FROM ((Employees INNER JOIN Orders ON Employees.EmployeeID = Orders.EmployeeID) INNER JOIN [Order Details] ON Orders.OrderID = [Order Details].OrderID) GROUP BY Employees.Country, DatePart(yyyy, Orders.OrderDate), DatePart(q, Orders.OrderDate)", "Data")
					ordersReport.AddTable(ds)
					Dim DataPath As String = Path.Combine(Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), ".."), "..") & Path.DirectorySeparatorChar

					If saveFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
						ordersReport.Run(DataPath & "Master Detail on one Table.template.xls", saveFileDialog1.FileName)

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
