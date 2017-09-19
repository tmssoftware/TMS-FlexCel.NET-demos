Imports System.Collections
Imports System.ComponentModel
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports System.IO
Imports System.Reflection
Imports System.Text

Imports FlexCel.Render

Namespace ExcelUserDefinedFunctions
	''' <summary>
	''' An example on how to recalculate user defined functions.
	''' </summary>
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private ReadOnly Property PathToExe() As String
			Get
				Return Path.Combine(Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), ".."), "..") & Path.DirectorySeparatorChar
			End Get
		End Property

		''' <summary>
		''' Loads the user defined functions into the Excel recalculating engine.
		''' </summary>
		''' <param name="Xls"></param>
		Private Sub LoadUdfs(ByVal Xls As ExcelFile)
			Xls.AddUserDefinedFunction(TUserDefinedFunctionScope.Local, TUserDefinedFunctionLocation.Internal, New SumCellsWithSameColor())
			Xls.AddUserDefinedFunction(TUserDefinedFunctionScope.Local, TUserDefinedFunctionLocation.Internal, New IsPrime())
			Xls.AddUserDefinedFunction(TUserDefinedFunctionScope.Local, TUserDefinedFunctionLocation.Internal, New BoolChoose())
			Xls.AddUserDefinedFunction(TUserDefinedFunctionScope.Local, TUserDefinedFunctionLocation.Internal, New Lowest())
		End Sub

		Private Sub AddData(ByVal Xls As ExcelFile)
			LoadUdfs(Xls) 'Register our custom functions. As we are using a local scope, we need to register them each time.

			Xls.Open(Path.Combine(PathToExe, "udfs.xls")) 'Open the file we want to manipulate.

			'Fill the cell range with other values so we can see how the sheet is recalculated by FlexCel.
			Dim Data As TXlsCellRange = Xls.GetNamedRange("Data", -1)
			For r As Integer = Data.Top To Data.Bottom - 1
				Xls.SetCellValue(r, Data.Left, r - Data.Top)
			Next r

			'Add an UDF to the sheet. We can enter the fucntion "BoolChoose" here because it was registered into FlexCel in LoadUDF()
			'If it hadn't been registered, this line would raise an Exception of an unknown function.
			Dim FmlaText As String = "=BoolChoose(TRUE,""This formula was entered with FlexCel!"",""It shouldn't display this"")"
			Xls.SetCellValue(11, 1, New TFormula(FmlaText))

			'Verify the UDF entered is correct. We can read any udf from Excel, even if it is not registered with AddUserDefinedFunction.
			Dim o As Object = Xls.GetCellValue(11, 1)
			Dim fm As TFormula = TryCast(o, TFormula)
			Debug.Assert(fm IsNot Nothing, "The cell must contain a formula")
			If fm IsNot Nothing Then
				Debug.Assert(fm.Text = FmlaText, "Error in Formula: It should be """ & FmlaText & """ and it is """ & fm.Text & """")
			End If

			'Recalc the sheet. As we are not saving it yet, we ned to make a manual recalc.
			Xls.Recalc()

			'Export the file to PDF so we can see the values calculated by FlexCel without Excel recalculating them.
			Using pdf As New FlexCelPdfExport(Xls, True)
				pdf.Export(saveFileDialog1.FileName)
			End Using

			'Save the file as xls too so we can compare.
			Xls.Save(Path.ChangeExtension(saveFileDialog1.FileName, "xls"))
		End Sub


		Private Sub button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles button1.Click
			AutoRun()
		End Sub

		Public Sub AutoRun()
			If saveFileDialog1.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
				Return
			End If
			Dim Xls As ExcelFile = New XlsFile(True)
			AddData(Xls)
			If MessageBox.Show("Do you want to open the generated files (PDF and XLS)?", "Confirm", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
				Process.Start(saveFileDialog1.FileName)
				Process.Start(Path.ChangeExtension(saveFileDialog1.FileName, "xls"))
			End If

		End Sub

		''' <summary>
		''' This is the method that will be called by the ASP.NET front end. It returns an array of bytes 
		''' with the report data, so the ASP.NET application can stream it to the client.
		''' </summary>
		''' <returns>The generated file as a byte array.</returns>
		Public Function WebRun() As Byte()
			Dim Xls As ExcelFile = New XlsFile(True)
			AddData(Xls)

			Using OutStream As New MemoryStream()
				Xls.Save(OutStream)
				Return OutStream.ToArray()
			End Using
		End Function
	End Class

	#Region "UDF definitions"
	''' <summary>
	''' Implements a custom function that will sum the cells in a range that have the same
	''' color of the source cell. This function mimics the VBA macro in the example, so when
	''' recalculating the sheet with FlexCel you will get the same results as with Excel.
	''' </summary>
	Public Class SumCellsWithSameColor
		Inherits TUserDefinedFunction

		''' <summary>
		''' Creates a new instance and registers the class in the FlexCel recalculating engine as "SumCellsWithSameColor".
		''' </summary>
		Public Sub New()
			MyBase.New("SumCellsWithSameColor")
		End Sub

		''' <summary>
		''' Returns the sum of cells in a range that have the same color as a reference cell.
		''' </summary>
		''' <param name="arguments"></param>
		''' <param name="parameters">In this case we expect 2 parameters, first the reference cell and then
		''' the range in which to sum. We will return an error otherwise.</param>
		''' <returns></returns>
		Public Overrides Function Evaluate(ByVal arguments As TUdfEventArgs, ByVal parameters() As Object) As Object
'			#Region "Get Parameters"
			Dim Err As TFlxFormulaErrorValue
			If Not CheckParameters(parameters, 2, Err) Then
				Return Err
			End If

			'The first parameter should be a range
			Dim SourceCell As TXls3DRange = Nothing
			If Not TryGetCellRange(parameters(0), SourceCell, Err) Then
				Return Err
			End If

			'The second parameter should be a range too.
			Dim SumRange As TXls3DRange = Nothing
			If Not TryGetCellRange(parameters(1), SumRange, Err) Then
				Return Err
			End If
'			#End Region

			'Get the color in SourceCell. Note that if Source cell is a range with more than one cell,
			'we will use the first cell in the range. Also, as different colors can have the same rgb value, we will compare the actual RGB values, not the ExcelColors
			Dim fmt As TFlxFormat = arguments.Xls.GetCellVisibleFormatDef(SourceCell.Sheet1, SourceCell.Top, SourceCell.Left)
			Dim SourceColor As Integer = fmt.FillPattern.FgColor.ToColor(arguments.Xls).ToArgb()

			Dim Result As Double = 0
			'Loop in the sum range and sum the corresponding values.
			For s As Integer = SumRange.Sheet1 To SumRange.Sheet2
				For r As Integer = SumRange.Top To SumRange.Bottom
					For c As Integer = SumRange.Left To SumRange.Right
						Dim XF As Integer = -1
						Dim val As Object = arguments.Xls.GetCellValue(s, r, c, XF)
						If TypeOf val Is Double Then 'we will only sum numeric values.
							Dim sumfmt As TFlxFormat = arguments.Xls.GetCellVisibleFormatDef(s, r, c)
							If sumfmt.FillPattern.FgColor.ToColor(arguments.Xls).ToArgb() = SourceColor Then
								Result += CDbl(val)
							End If
						End If
					Next c
				Next r
			Next s
			Return Result
		End Function
	End Class


	''' <summary>
	''' Implements a custom function that will return true if a number is prime.
	''' This function mimics the VBA macro in the example, so when
	''' recalculating the sheet with FlexCel you will get the same results as with Excel.
	''' </summary>
	Public Class IsPrime
		Inherits TUserDefinedFunction

		''' <summary>
		''' Creates a new instance and registers the class in the FlexCel recalculating engine as "IsPrime".
		''' </summary>
		Public Sub New()
			MyBase.New("IsPrime")
		End Sub

		''' <summary>
		''' Returns true if a number is prime.
		''' </summary>
		''' <param name="arguments"></param>
		''' <param name="parameters">In this case we expect 1 parameter with the number. We will return an error otherwise.</param>
		''' <returns></returns>
		Public Overrides Function Evaluate(ByVal arguments As TUdfEventArgs, ByVal parameters() As Object) As Object
'			#Region "Get Parameters"
			Dim Err As TFlxFormulaErrorValue
			If Not CheckParameters(parameters, 1, Err) Then
				Return Err
			End If

			'The parameter should be a double or a range.
			Dim Number As Double
			If Not TryGetDouble(arguments.Xls, parameters(0), Number, Err) Then
				Return Err
			End If
'			#End Region

			'Return true if the number is prime.
			Dim n As Integer = Convert.ToInt32(Number)
			If n = 2 Then
				Return True
			End If
			If n < 2 OrElse n Mod 2 = 0 Then
				Return False
			End If
			For i As Integer = 3 To Convert.ToInt32(Fix(Math.Sqrt(n))) Step 2
				If n Mod i = 0 Then
					Return False
				End If
			Next i
			Return True
		End Function
	End Class

	''' <summary>
	''' Implements a custom function that will choose between two different strings.
	''' This function mimics the VBA macro in the example, so when
	''' recalculating the sheet with FlexCel you will get the same results as with Excel.
	''' </summary>
	Public Class BoolChoose
		Inherits TUserDefinedFunction

		''' <summary>
		''' Creates a new instance and registers the class in the FlexCel recalculating engine as "BoolChoose".
		''' </summary>
		Public Sub New()
			MyBase.New("BoolChoose")
		End Sub

		''' <summary>
		''' Chooses between 2 different strings.
		''' </summary>
		''' <param name="arguments"></param>
		''' <param name="parameters">In this case we expect 3 parameters: The first is a boolean, and the other 2 strings. We will return an error otherwise.</param>
		''' <returns></returns>
		Public Overrides Function Evaluate(ByVal arguments As TUdfEventArgs, ByVal parameters() As Object) As Object
'			#Region "Get Parameters"
			Dim Err As TFlxFormulaErrorValue
			If Not CheckParameters(parameters, 3, Err) Then
				Return Err
			End If

			'The first parameter should be a boolean.
			Dim ChooseFirst As Boolean
			If Not TryGetBoolean(arguments.Xls, parameters(0), ChooseFirst, Err) Then
				Return Err
			End If

			'The second parameter should be a string.
			Dim s1 As String = Nothing
			If Not TryGetString(arguments.Xls, parameters(1), s1, Err) Then
				Return Err
			End If

			'The third parameter should be a string.
			Dim s2 As String = Nothing
			If Not TryGetString(arguments.Xls, parameters(2), s2, Err) Then
				Return Err
			End If
'			#End Region

			'Return s1 or s2 depending on ChooseFirst
			If ChooseFirst Then
				Return s1
			Else
				Return s2
			End If
		End Function
	End Class

	''' <summary>
	''' Implements a custom function that will choose the lowest member in an array.
	''' This function mimics the VBA macro in the example, so when
	''' recalculating the sheet with FlexCel you will get the same results as with Excel.
	''' </summary>
	Public Class Lowest
		Inherits TUserDefinedFunction

		''' <summary>
		''' Creates a new instance and registers the class in the FlexCel recalculating engine as "Lowest".
		''' </summary>
		Public Sub New()
			MyBase.New("Lowest")
		End Sub

		''' <summary>
		''' Chooses the lowest element in an array.
		''' </summary>
		''' <param name="arguments"></param>
		''' <param name="parameters">In this case we expect 1 parameter that should be an array. We will return an error otherwise.</param>
		''' <returns></returns>
		Public Overrides Function Evaluate(ByVal arguments As TUdfEventArgs, ByVal parameters() As Object) As Object
'			#Region "Get Parameters"
			Dim Err As TFlxFormulaErrorValue
			If Not CheckParameters(parameters, 1, Err) Then
				Return Err
			End If

			'The first parameter should be an array.
			Dim SourceArray(,) As Object = Nothing
			If Not TryGetArray(arguments.Xls, parameters(0), SourceArray, Err) Then
				Return Err
			End If
'			#End Region

			Dim Result As Double = 0
			Dim First As Boolean = True
			For Each o As Object In SourceArray
				If TypeOf o Is Double Then
					If First Then
						First = False
						Result = CDbl(o)
					Else
						If CDbl(o) < Result Then
							Result = CDbl(o)
						End If
					End If
				Else
					Return TFlxFormulaErrorValue.ErrValue
				End If
			Next o

			Return Result
		End Function

	End Class
	#End Region

End Namespace
