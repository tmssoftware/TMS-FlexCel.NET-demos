Imports System.Collections
Imports System.ComponentModel
Imports System.Reflection
Imports FlexCel.Core

Namespace MainDemo
	''' <summary>
	''' About...
	''' </summary>
	Partial Public Class AboutForm
		Inherits System.Windows.Forms.Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles button1.Click
			Close()
		End Sub

		Private Sub linkLabel1_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles linkLabel1.LinkClicked
			System.Diagnostics.Process.Start(linkLabel1.Text)
		End Sub

		Private Sub linkLabel2_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles linkLabel2.LinkClicked
			System.Diagnostics.Process.Start(linkLabel2.Text)
		End Sub

		Private Sub AboutForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
			Dim asm As System.Reflection.Assembly = System.Reflection.Assembly.GetAssembly(GetType(ExcelFile))
			lblVersion.Text = "Using FlexCel Version: " & asm.GetName().Version.ToString()

		End Sub
	End Class
End Namespace
