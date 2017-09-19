Imports System.Collections
Imports System.ComponentModel

Namespace MainDemo
	''' <summary>
	''' Summary description for DbErrorForm.
	''' </summary>
	Partial Public Class DbErrorForm
		Inherits System.Windows.Forms.Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub LinkToDownload_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkToDownload.LinkClicked
			System.Diagnostics.Process.Start(LinkToDownload.Text)
		End Sub

		Private Sub btnCopy2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCopy.Click
			Clipboard.SetDataObject(LinkToDownload.Text)
		End Sub
	End Class
End Namespace
