Imports System.ComponentModel
Imports System.Text

Namespace VirtualMode
	Partial Public Class SheetSelectorForm
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Public Sub New(ByVal SheetNames() As String)
			Me.New()
			For Each s As String In SheetNames
				SheetList.Items.Add(s)
			Next s

			SheetList.SelectedIndex = 0
		End Sub

		Friend Function Execute() As Boolean
			Return ShowDialog() = System.Windows.Forms.DialogResult.OK
		End Function

		Public ReadOnly Property SelectedSheet() As String
			Get
				Return Convert.ToString(SheetList.SelectedItem)
			End Get
		End Property

		Public ReadOnly Property SelectedSheetIndex() As Integer
			Get
				Return SheetList.SelectedIndex
			End Get
		End Property

		Private Sub SheetList_DoubleClick(ByVal sender As Object, ByVal e As EventArgs) Handles SheetList.DoubleClick
			DialogResult = System.Windows.Forms.DialogResult.OK
		End Sub
	End Class
End Namespace
