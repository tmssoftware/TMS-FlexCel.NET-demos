Imports System.Collections
Imports System.ComponentModel

Namespace EncryptedFiles
	''' <summary>
	''' Summary description for PasswordDialog.
	''' </summary>
	Partial Public Class PasswordDialog
		Inherits System.Windows.Forms.Form

		Public Sub New()
			'
			' Required for Windows Form Designer support
			'
			InitializeComponent()

			'
			' TODO: Add any constructor code after InitializeComponent call
			'
		End Sub

		Public ReadOnly Property Password() As String
			Get
				Return PasswordEdit.Text
			End Get
		End Property
	End Class
End Namespace
