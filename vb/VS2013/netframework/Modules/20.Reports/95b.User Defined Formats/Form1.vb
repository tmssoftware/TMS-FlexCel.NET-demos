Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Report
Imports FlexCel.Demo.SharedData


Namespace UserDefinedFormats
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles button1.Click
			AutoRun()
		End Sub

		Public Sub AutoRun()
			Using ordersReport As FlexCelReport = SharedData.CreateReport()
				ordersReport.SetValue("Date", Date.Now)
				ordersReport.SetUserFormat("ZipCode", New ZipCodeImp())
				ordersReport.SetUserFormat("ShipFormat", New ShipFormatImp())

				Dim DataPath As String = Path.Combine(Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), ".."), "..") & Path.DirectorySeparatorChar

				If saveFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
					ordersReport.Run(DataPath & "User Defined Formats.template.xlsx", saveFileDialog1.FileName)

					If MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
						Process.Start(saveFileDialog1.FileName)
					End If
				End If
			End Using
		End Sub


		Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
			Close()
		End Sub

	End Class

	#Region "ZipCode Implementation"
	Friend Class ZipCodeImp
		Inherits TFlexCelUserFormat

		Public Sub New()
		End Sub

		Public Overrides Function Evaluate(ByVal workbook As ExcelFile, ByVal rangeToFormat As TXlsCellRange, ByVal parameters() As Object) As TFlxPartialFormat
			If parameters Is Nothing OrElse parameters.Length <> 1 Then
				Throw New ArgumentException("Bad parameter count in call to ZipCode() user-defined format")
			End If

			Dim color As Integer
			'If the zip code is not valid, don't modify the format.
			If parameters(0) Is Nothing OrElse (Not Integer.TryParse(Convert.ToString(parameters(0)), color)) Then
				Return New TFlxPartialFormat(Nothing, Nothing, False)
			End If

			'This code is not supposed to make sense. We will convert the zip code to a color based in the numeric value.
			Dim fmt As TFlxFormat = workbook.GetDefaultFormat
			fmt.FillPattern.Pattern = TFlxPatternStyle.Solid
			fmt.FillPattern.FgColor = TExcelColor.FromArgb(color)
			fmt.FillPattern.BgColor = TExcelColor.Automatic

			fmt.Font.Color = TExcelColor.FromArgb((Not color))

			Dim apply As New TFlxApplyFormat()
			apply.FillPattern.SetAllMembers(True)
			apply.Font.Color = True
			Return New TFlxPartialFormat(fmt, apply, False)
		End Function
	End Class
	#End Region

	#Region "ShipFormat Implementation"
	Friend Class ShipFormatImp
		Inherits TFlexCelUserFormat

		Public Sub New()
		End Sub

		Public Overrides Function Evaluate(ByVal workbook As ExcelFile, ByVal rangeToFormat As TXlsCellRange, ByVal parameters() As Object) As TFlxPartialFormat
			'Again, this example is not supposed to make sense, only to show how you can code a complex rule.
			'This method will format the rows with a color that depends in the length of the first parameter,
			'and if the second parameter starts with "B" it will make the text red.

			If parameters Is Nothing OrElse parameters.Length <> 2 Then
				Throw New ArgumentException("Bad parameter count in call to ShipFormat() user-defined format")
			End If

			Dim len As Integer = Convert.ToString(parameters(0)).Length
			Dim country As String = Convert.ToString(parameters(1))

			Dim color As Int32 = &HFFFFFF - len * 100
			Dim fmt As TFlxFormat = workbook.GetDefaultFormat
			fmt.FillPattern.Pattern = TFlxPatternStyle.Solid
			fmt.FillPattern.FgColor = TExcelColor.FromArgb(color)
			fmt.FillPattern.BgColor = TExcelColor.Automatic

			Dim apply As New TFlxApplyFormat()
			apply.FillPattern.SetAllMembers(True)

			If country.StartsWith("B") Then
				fmt.Font.Color = Colors.OrangeRed
				apply.Font.Color = True
			End If

			Return New TFlxPartialFormat(fmt, apply, False)
		End Function
	End Class
	#End Region

End Namespace
