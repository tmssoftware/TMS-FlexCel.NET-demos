Imports System.ComponentModel

Namespace CustomPreview
	''' <summary>
	''' Form for asking for a password when the file is password protected.
	''' </summary>
	Partial Public Class PasswordForm
		Inherits System.Windows.Forms.Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Public ReadOnly Property Password() As String
			Get
				Return PasswordEdit.Text
			End Get
		End Property
	End Class
End Namespace
