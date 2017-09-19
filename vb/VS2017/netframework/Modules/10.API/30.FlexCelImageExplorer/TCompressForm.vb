Imports System.Collections
Imports System.ComponentModel
Imports System.IO

Namespace FlexCelImageExplorer
	''' <summary>
	''' Summary description for TCompressForm.
	''' </summary>
	Partial Public Class TCompressForm
		Inherits System.Windows.Forms.Form

		Public Sub New()
			'
			' Required for Windows Form Designer support
			'
			InitializeComponent()
			cbPixelFormat.SelectedIndex = 2

			'
			' TODO: Add any constructor code after InitializeComponent call
			'
		End Sub

		Private FImageToUse() As Byte
		Private FXlsFilename As String

		Friend Property ImageToUse() As Byte()
			Get
				Return FImageToUse
			End Get
			Set(ByVal value As Byte())
				FImageToUse = value
				Using ms As New MemoryStream(value)
					pictureBox1.Image = Image.FromStream(ms)
				End Using
			End Set
		End Property

		Friend Property XlsFilename() As String
			Get
				Return FXlsFilename
			End Get
			Set(ByVal value As String)
				FXlsFilename = value
			End Set
		End Property

		Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
			Close()
		End Sub

		Private Sub TCompressForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
		End Sub

		Private Sub btnOk_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOk.Click
			Close()
		End Sub
	End Class
End Namespace
