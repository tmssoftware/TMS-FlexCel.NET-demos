Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Report
Imports FlexCel.Demo.SharedData


Namespace EncryptionAndSubtotals
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles button1.Click
			ManualRun()
		End Sub

		Public Sub ManualRun()
			Using ordersReport As FlexCelReport = SharedData.CreateReport()
				AddHandler ordersReport.BeforeReadTemplate, AddressOf ordersReport_BeforeReadTemplate
				AddHandler ordersReport.AfterGenerateSheet, AddressOf ordersReport_AfterGenerateSheet
				AddHandler ordersReport.AfterGenerateWorkbook, AddressOf ordersReport_AfterGenerateWorkbook

				ordersReport.SetValue("Date", Date.Now)

				Dim DataPath As String = Path.Combine(Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), ".."), "..") & Path.DirectorySeparatorChar

				If saveFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
					ordersReport.Run(DataPath & "Encryption And Subtotals.template.xls", saveFileDialog1.FileName)

					If MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
						Process.Start(saveFileDialog1.FileName)
					End If
				End If
			End Using
		End Sub

		Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
			Close()
		End Sub

		Private Sub mainForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
			encryptionType.SelectedItem = encryptionType.Items(0)
		End Sub


		Private Sub ordersReport_BeforeReadTemplate(ByVal sender As Object, ByVal e As FlexCel.Report.GenerateEventArgs)
			e.File.Protection.OpenPassword = OpenPassTemplate.Text
		End Sub

		Private Sub ordersReport_AfterGenerateSheet(ByVal sender As Object, ByVal e As FlexCel.Report.GenerateEventArgs)
			e.File.Protection.SetSheetProtection(ProtectSheetPass.Text, New TSheetProtectionOptions(ProtectSheet.Checked))
		End Sub

		Private Sub ordersReport_AfterGenerateWorkbook(ByVal sender As Object, ByVal e As FlexCel.Report.GenerateEventArgs)
			If encryptionType.SelectedItem Is encryptionType.Items(1) Then
				e.File.Protection.EncryptionType = TEncryptionType.Xor
			Else
				e.File.Protection.EncryptionType = TEncryptionType.Standard
			End If
			e.File.Protection.OpenPassword = OpenPassGenerated.Text
			e.File.Protection.SetModifyPassword(ModifyPassGenerated.Text, RecommendReadOnly.Checked, ReservingUser.Text)
			e.File.Protection.SetWorkbookProtection(ProtectWorkbookPass.Text, New TWorkbookProtectionOptions(False, ProtectWorkbook.Checked))
		End Sub
	End Class

End Namespace
