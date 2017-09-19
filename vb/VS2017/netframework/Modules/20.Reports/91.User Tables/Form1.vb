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


Namespace UserTables
	''' <summary>
	''' Using tables that are defined in the template.
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
			Using genericReport As New FlexCelReport(True)
				AddHandler genericReport.UserTable, AddressOf genericReport_UserTable
				genericReport.DeleteEmptyRanges = False

				Dim DataPath As String = Path.Combine(Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), ".."), "..") & Path.DirectorySeparatorChar

				If saveFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
					genericReport.Run(DataPath & "User Tables.template" & Path.GetExtension(saveFileDialog1.FileName), saveFileDialog1.FileName)

					If MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
						Process.Start(saveFileDialog1.FileName)
					End If
				End If
			End Using
		End Sub

		Private Sub genericReport_UserTable(ByVal sender As Object, ByVal e As UserTableEventArgs)
			Dim ds As New DataSet()

			'On this example we will just return the table with the name specified on parameters
			'but you could return whatever you wanted here.
			'As always, remember to *validate* what the user can enter on the parameters string.

			Select Case e.Parameters.ToUpper(CultureInfo.InvariantCulture)
				Case "SUPPLIERS"
					SharedData.Fill(ds, "select * from suppliers", e.TableName)
				Case "CATEGORIES"
					SharedData.Fill(ds, "select * from categories", e.TableName)
				Case "PRODUCTS"
					SharedData.Fill(ds, "select * from products", e.TableName)

				Case Else
					Throw New Exception("Invalid parameter to user table: " & e.Parameters)
			End Select

			CType(sender, FlexCelReport).AddTable(ds, TDisposeMode.DisposeAfterRun)
		End Sub

		Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
			Close()
		End Sub
	End Class


End Namespace
