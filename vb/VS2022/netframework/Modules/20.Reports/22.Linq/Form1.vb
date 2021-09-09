Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Report


Namespace Linq
	''' <summary>
	''' Summary description for Form1.
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
			Using report As New FlexCelReport(True)
				LoadTables(report)

				Dim DataPath As String = Path.Combine(Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), ".."), "..") & Path.DirectorySeparatorChar

				If saveFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
					report.Run(DataPath & "Linq.template.xls", saveFileDialog1.FileName)

					If MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
						Process.Start(saveFileDialog1.FileName)
					End If
				End If
			End Using
		End Sub

		Private Sub LoadTables(ByVal report As FlexCelReport)
			Dim Categories As New List(Of Categories)()
			Dim Animals As New Categories("Animals")
			Animals.Elements.Add(New Elements(1, "Penguin"))
			Animals.Elements.Add(New Elements(2, "Cat"))
			Animals.Elements.Add(New Elements(3, "Unicorn"))
			Categories.Add(Animals)

			Dim Flowers As New Categories("Flowers")
			Flowers.Elements.Add(New Elements(4, "Daisy"))
			Flowers.Elements.Add(New Elements(5, "Rose"))
			Flowers.Elements.Add(New Elements(6, "Orchid"))
			Categories.Add(Flowers)

			report.AddTable("Categories", Categories)
			'We don't need to call AddTable for elements since it is already added when we add Categories.


			Dim ElementNames As New List(Of ElementName)()
			ElementNames.Add(New ElementName(1, "Linus"))
			ElementNames.Add(New ElementName(1, "Gerard"))
			ElementNames.Add(New ElementName(2, "Rover"))
			ElementNames.Add(New ElementName(3, "Mike"))
			ElementNames.Add(New ElementName(5, "Rosalyn"))
			ElementNames.Add(New ElementName(5, "Monica"))
			ElementNames.Add(New ElementName(6, "Lisa"))

			report.AddTable("ElementName", ElementNames)
			'ElementName doesn't have an intrinsic relationship with categories, so we will have to manually add a relationship.
			'Non intrinsic relationships should be rare, but we do it here to show how it can be done.
			report.AddRelationship("Elements", "ElementName", "ElementID", "ElementID")
		End Sub

		Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
			Close()
		End Sub
	End Class

	Public Class Categories
		'Public properties can be used in reports.
		Private privateName As String
		Public Property Name() As String
			Get
				Return privateName
			End Get
			Private Set(ByVal value As String)
				privateName = value
			End Set
		End Property

		'Elements is in master-detail relationship with this element, even when we don't explicitly add a relationship. 
		'Relationship is inferred because Elements is a property of this object
		Private privateElements As List(Of Elements)
		Public Property Elements() As List(Of Elements)
			Get
				Return privateElements
			End Get
			Private Set(ByVal value As List(Of Elements))
				privateElements = value
			End Set
		End Property

		Public Sub New(ByVal name As String)
			Me.Name = name
			Elements = New List(Of Elements)()
		End Sub

	End Class

	Public Class Elements
		'We will relate this property with the table of colors by adding a relationship.
		Private privateElementID As Integer
		Public Property ElementID() As Integer
			Get
				Return privateElementID
			End Get
			Private Set(ByVal value As Integer)
				privateElementID = value
			End Set
		End Property

		Private privateName As String
		Public Property Name() As String
			Get
				Return privateName
			End Get
			Private Set(ByVal value As String)
				privateName = value
			End Set
		End Property


		Public Sub New(ByVal elementID As Integer, ByVal name As String)
			Me.Name = name
			Me.ElementID = elementID
		End Sub

	End Class

	Public Class ElementName
		Private privateElementID As Integer
		Public Property ElementID() As Integer
			Get
				Return privateElementID
			End Get
			Private Set(ByVal value As Integer)
				privateElementID = value
			End Set
		End Property
		Private privateName As String
		Public Property Name() As String
			Get
				Return privateName
			End Get
			Private Set(ByVal value As String)
				privateName = value
			End Set
		End Property

		Public Sub New(ByVal elementID As Integer, ByVal name As String)
			Me.ElementID = elementID
			Me.Name = name
		End Sub
	End Class

End Namespace
