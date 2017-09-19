Imports System.Collections
Imports System.ComponentModel

Namespace GenericReports
	''' <summary>
	''' A dialog where you can enter any SQL.
	''' </summary>
	Partial Public Class EnterSQLDialog
		Inherits System.Windows.Forms.Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Public ReadOnly Property SQL() As String
			Get
				Return edSQL.Text
			End Get
		End Property
	End Class
End Namespace
