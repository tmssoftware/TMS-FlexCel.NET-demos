Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Report


Namespace AdvancedLinq
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
					report.Run(DataPath & "Advanced Linq.template.xlsx", saveFileDialog1.FileName)

					If MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
						Process.Start(saveFileDialog1.FileName)
					End If
				End If
			End Using
		End Sub

		Private Sub LoadTables(ByVal report As FlexCelReport)
			Dim Countries = New List(Of Country)()
			Countries.Add(New Country("China", New People(1384688986), New Geography(New Area(270550, 9326410))))

			Dim country = Countries(Countries.Count - 1)
			country.People.Language.Add(New Language(New LanguageName("Md", "Mandarin"), New LanguageSpeakers(0, 66.2)))

			country.People.Language.Add(New Language(New LanguageName("Yue", "Yue"), New LanguageSpeakers(0, 4.9)))

			country.People.Language.Add(New Language(New LanguageName("Wu", "Wu"), New LanguageSpeakers(0, 6.1)))

			country.People.Language.Add(New Language(New LanguageName("Mb", "Minbei"), New LanguageSpeakers(0, 6.2)))

			country.People.Language.Add(New Language(New LanguageName("Mn", "Minnan"), New LanguageSpeakers(0, 5.2)))

			country.People.Language.Add(New Language(New LanguageName("Xi", "Xiang"), New LanguageSpeakers(0, 3.0)))

			country.People.Language.Add(New Language(New LanguageName("Gan", "Gan"), New LanguageSpeakers(0, 4.0)))


			Countries.Add(New Country("India", New People(1296834042), New Geography(New Area(314070, 2973193))))

			country = Countries(Countries.Count - 1)
			country.People.Language.Add(New Language(New LanguageName("Hi", "Hindi"), New LanguageSpeakers(0, 43.6)))

			country.People.Language.Add(New Language(New LanguageName("Bg", "Bengali"), New LanguageSpeakers(0, 8)))

			country.People.Language.Add(New Language(New LanguageName("Ma", "Marath"), New LanguageSpeakers(0, 6.9)))

			country.People.Language.Add(New Language(New LanguageName("Te", "Telugu"), New LanguageSpeakers(0, 6.7)))

			country.People.Language.Add(New Language(New LanguageName("Ta", "Tamil"), New LanguageSpeakers(0, 5.7)))

			country.People.Language.Add(New Language(New LanguageName("Gu", "Gujarati"), New LanguageSpeakers(0, 4.6)))

			country.People.Language.Add(New Language(New LanguageName("Ur", "Urdu"), New LanguageSpeakers(0, 4.2)))

			country.People.Language.Add(New Language(New LanguageName("Ka", "Kannada"), New LanguageSpeakers(0, 3.6)))

			country.People.Language.Add(New Language(New LanguageName("Od", "Odia"), New LanguageSpeakers(0, 3.1)))

			country.People.Language.Add(New Language(New LanguageName("Ma", "Malayalam"), New LanguageSpeakers(0, 2.9)))

			country.People.Language.Add(New Language(New LanguageName("Pu", "Punjabi"), New LanguageSpeakers(0, 2.7)))

			country.People.Language.Add(New Language(New LanguageName("As", "Assamese"), New LanguageSpeakers(0, 1.3)))

			country.People.Language.Add(New Language(New LanguageName("Mi", "Maithili"), New LanguageSpeakers(0, 1.1)))

			country.People.Language.Add(New Language(New LanguageName("O", "Other"), New LanguageSpeakers(0, 5.6)))


			Countries.Add(New Country("United States", New People(329256465), New Geography(New Area(685924, 9147593))))

			country = Countries(Countries.Count - 1)
			country.People.Language.Add(New Language(New LanguageName("En", "English"), New LanguageSpeakers(0, 78.2)))

			country.People.Language.Add(New Language(New LanguageName("Sp", "Spanish"), New LanguageSpeakers(0, 13.4)))

			country.People.Language.Add(New Language(New LanguageName("Ch", "Chinese"), New LanguageSpeakers(0, 1.1)))

			country.People.Language.Add(New Language(New LanguageName("O", "Other"), New LanguageSpeakers(0, 7.3)))

			report.AddTable("country", Countries)
		End Sub

		Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
			Close()
		End Sub
	End Class

	Public Class Country
		Private privateName As String
		Public Property Name() As String
			Get
				Return privateName
			End Get
			Private Set(ByVal value As String)
				privateName = value
			End Set
		End Property

		Public Property People() As People
		Public Property Geography() As Geography

		Public Sub New(ByVal name As String, ByVal people As People, ByVal geography As Geography)
			Me.Name = name
			Me.People = people
			Me.Geography = geography
		End Sub

	End Class

	Public Class Geography
		Private privateArea As Area
		Public Property Area() As Area
			Get
				Return privateArea
			End Get
			Private Set(ByVal value As Area)
				privateArea = value
			End Set
		End Property

		Public Sub New(ByVal area As Area)
			Me.Area = area
		End Sub
	End Class

	Public Class Area
		Public ReadOnly Property Total() As Integer
			Get
				Return Water + Land
			End Get
		End Property
		Private privateWater As Integer
		Public Property Water() As Integer
			Get
				Return privateWater
			End Get
			Private Set(ByVal value As Integer)
				privateWater = value
			End Set
		End Property
		Private privateLand As Integer
		Public Property Land() As Integer
			Get
				Return privateLand
			End Get
			Private Set(ByVal value As Integer)
				privateLand = value
			End Set
		End Property

		Public Sub New(ByVal water As Integer, ByVal land As Integer)
			Me.Water = water
			Me.Land = land
		End Sub
	End Class

	Public Class People
		Private privatePopulation As Integer
		Public Property Population() As Integer
			Get
				Return privatePopulation
			End Get
			Private Set(ByVal value As Integer)
				privatePopulation = value
			End Set
		End Property
		Private privateLanguage As List(Of Language)
		Public Property Language() As List(Of Language)
			Get
				Return privateLanguage
			End Get
			Private Set(ByVal value As List(Of Language))
				privateLanguage = value
			End Set
		End Property

		Public Sub New(ByVal population As Integer)
			Me.Population = population
			Language = New List(Of Language)()
		End Sub
	End Class

	Public Class Language
		Private privateName As LanguageName
		Public Property Name() As LanguageName
			Get
				Return privateName
			End Get
			Private Set(ByVal value As LanguageName)
				privateName = value
			End Set
		End Property
		Private privateSpeakers As LanguageSpeakers
		Public Property Speakers() As LanguageSpeakers
			Get
				Return privateSpeakers
			End Get
			Private Set(ByVal value As LanguageSpeakers)
				privateSpeakers = value
			End Set
		End Property

		Public Sub New(ByVal name As LanguageName, ByVal speakers As LanguageSpeakers)
			Me.Name = name
			Me.Speakers = speakers
		End Sub

	End Class

	Public Class LanguageName
		Private privateShortName As String
		Public Property ShortName() As String
			Get
				Return privateShortName
			End Get
			Private Set(ByVal value As String)
				privateShortName = value
			End Set
		End Property
		Private privateLongName As String
		Public Property LongName() As String
			Get
				Return privateLongName
			End Get
			Private Set(ByVal value As String)
				privateLongName = value
			End Set
		End Property

		Public Sub New(ByVal shortName As String, ByVal longName As String)
			Me.ShortName = shortName
			Me.LongName = longName
		End Sub
	End Class

	Public Class LanguageSpeakers
		Private privateAbsoluteNumber As Integer
		Public Property AbsoluteNumber() As Integer
			Get
				Return privateAbsoluteNumber
			End Get
			Private Set(ByVal value As Integer)
				privateAbsoluteNumber = value
			End Set
		End Property
		Private privatePercent As Double
		Public Property Percent() As Double
			Get
				Return privatePercent
			End Get
			Private Set(ByVal value As Double)
				privatePercent = value
			End Set
		End Property

		Public Sub New(ByVal absoluteNumber As Integer, ByVal percent As Double)
			Me.AbsoluteNumber = absoluteNumber
			Me.Percent = percent / 100.0
		End Sub
	End Class
End Namespace
