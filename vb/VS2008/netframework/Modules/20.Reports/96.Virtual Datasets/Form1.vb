Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports System.Resources
Imports System.Globalization
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Report


Namespace VirtualDatasets
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles button1.Click
			AutoRun()
		End Sub

		Public Sub AutoRun()
			Dim DataPath As String = Path.Combine(Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), ".."), "..") & Path.DirectorySeparatorChar

			If saveFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
				Dim SimpleData()() As Object = LoadDataSet(Path.Combine(DataPath, "Countries.txt"))
				Dim SimpleTable As New SimpleVirtualArrayDataSource(Nothing, SimpleData, New String() { "Rank", "Country", "Area", "Date" }, "SimpleTable")

				Using genericReport As New FlexCelReport(True)
					genericReport.AddTable("SimpleData", SimpleTable)

					Dim Complex1()() As Object = LoadDataSet(Path.Combine(DataPath, "Countries.txt"))
					Dim ComplexAreas As New ComplexVirtualArrayDataSource(Nothing, Complex1, New String() { "Rank", "Country", "Area", "Date" }, "ComplexAreas")
					Dim Complex2()() As Object = LoadDataSet(Path.Combine(DataPath, "Populations.txt"))
					Dim ComplexPopulations As New ComplexVirtualArrayDataSource(Nothing, Complex2, New String() { "Rank", "Country", "Population", "Date" }, "ComplexPopulations")

					genericReport.AddTable("ComplexAreas", ComplexAreas, TDisposeMode.DisposeAfterRun)
					genericReport.AddTable("ComplexPopulations", ComplexPopulations, TDisposeMode.DisposeAfterRun)



					genericReport.Run(Path.Combine(DataPath, "Virtual Datasets.template.xls"), saveFileDialog1.FileName)
				End Using

				If MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
					Process.Start(saveFileDialog1.FileName)
				End If
			End If
		End Sub


		Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
			Close()
		End Sub

		Private Function LoadDataSet(ByVal filename As String) As Object()()
			'Let's create some bussiness object with random data.

			Dim Result As New ArrayList()
			Using sr As New StreamReader(Path.GetFullPath(filename))
				Dim line As String
				line = sr.ReadLine()
				Do While line IsNot Nothing
					Dim fields() As String = line.Split(ControlChars.Tab)
					'Zero validation here since this is a demo and will use always the same data. On a real app you should not expect your data to play nice
					Dim f(fields.Length - 1) As Object
					Dim s As String = TryCast(fields(0), String)
					f(0) = Convert.ToInt64(s)
					f(1) = fields(1)
					s = TryCast(fields(2), String)
					f(2) = CObj(Convert.ToInt64(s.Replace(",", "")))
					f(3) = fields(3)
					Result.Add(f)
					line = sr.ReadLine()
				Loop
			End Using

			Return CType(Result.ToArray(GetType(Object())), Object()())
		End Function
	End Class

End Namespace
