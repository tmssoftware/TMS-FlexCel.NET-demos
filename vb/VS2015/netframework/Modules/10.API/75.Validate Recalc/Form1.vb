Imports System.Collections
Imports System.ComponentModel
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports System.IO
Imports System.Reflection

Namespace ValidateRecalc
	''' <summary>
	''' Use this demo to validate the recalculation made by FlexCel.
	''' </summary>
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Private XlsReport As FlexCel.Report.FlexCelReport

		Public Sub New()
			InitializeComponent()
			ResizeToolbar(mainToolbar)
		End Sub

		Private Sub ResizeToolbar(ByVal toolbar As ToolStrip)

			Using gr As Graphics = CreateGraphics()
				Dim xFactor As Double = gr.DpiX / 96.0
				Dim yFactor As Double = gr.DpiY / 96.0
				toolbar.ImageScalingSize = New Size(CInt(Fix(24 * xFactor)), CInt(Fix(24 * yFactor)))
				toolbar.Width = 0 'force a recalc of the buttons.
			End Using
		End Sub


		Private Sub button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.Click
			Close()
		End Sub

		Private Sub btnInfo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnInfo.Click
			MessageBox.Show("This example will validate the calculations performed by the FlexCel engine." & vbLf & "It can do it in 2 different ways:" & vbLf & "  1) The button 'Validate Recalc' will analyze a file, and report if there is anything that FlexCel doesn't support on it." & vbLf & "  2) The button 'Compare with Excel' will open a file saved by Excel, recalculate it with FlexCel, compare the values reported by both FlexCel and Excel and report if there are any differences.")
		End Sub

		Private Sub validateRecalc_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles validateRecalc.Click
			If openFileDialog1.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
				Return
			End If
			Dim Xls As New XlsFile()

			Xls.Open(openFileDialog1.FileName)

			' /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			' ////////Code here is only needed if you have linked files. In this example we don't know, so we will use it /////////
			Dim Work As New TWorkspace() 'Create a workspace
			Work.Add(Path.GetFileName(openFileDialog1.FileName), Xls) 'Add the original file to it
			AddHandler Work.LoadLinkedFile, AddressOf Work_LoadLinkedFile 'Set up an event to load the linked files.
																						 ' /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

			report.Text = "Results on file: " & openFileDialog1.FileName
			Dim Usl As TUnsupportedFormulaList = Xls.RecalcAndVerify()
			If Usl.Count = 0 Then
				report.Text &= vbLf & "**********All formulas supported!**********"
				Return
			End If

			report.Text &= vbLf & "Issues Found:"
			For i As Integer = 0 To Usl.Count - 1
				Dim FileName As String = String.Empty
				If Usl(i).FileName IsNot Nothing Then
					FileName = "File: " & Usl(i).FileName & "  => "
				End If
				report.Text &= vbLf & "     " & FileName & Usl(i).Cell.CellRef & ": " & Usl(i).ErrorType.ToString()
				If Usl(i).StackTrace IsNot Nothing Then
					For k As Integer = 0 To Usl(i).StackTrace.Length - 1
						If Usl(i).StackTrace(k).Address IsNot Nothing Then
							Dim TraceFileName As String = ""
							If Usl(i).StackTrace(k).FileName Is Nothing Then TraceFileName = "[" & Usl(i).StackTrace(k).FileName & "]"
							report.Text &= vbLf & "         -> References cell: " & TraceFileName & Usl(i).StackTrace(k).Address.CellRef
						End If
					Next k
				End If
				If Usl(i).FunctionName IsNot Nothing Then
					Dim FunctionStr As String = "Function"
					If Usl(i).ErrorType = TUnsupportedFormulaErrorType.ExternalReference Then
						FunctionStr = "Linked file not found"
					End If
					report.Text &= " ->" & FunctionStr & ": " & Usl(i).FunctionName
				End If
			Next i
		End Sub

		Private Sub compareWithExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles compareWithExcel.Click
			If openFileDialog1.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
				Return
			End If
			compareWithExcel.Enabled = False
			validateRecalc.Enabled = False
			Try
				Dim xls1 As New XlsFile()
				Dim xls2 As New XlsFile()

				xls1.Open(openFileDialog1.FileName)
				xls2.Open(openFileDialog1.FileName)
				report.Text = "Compare with Excel: " & openFileDialog1.FileName

				' /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
				' ////////Code here is only needed if you have linked files. In this example we don't know, so we will use it /////////
				Dim Work As New TWorkspace() 'Create a workspace
				Work.Add(Path.GetFileName(openFileDialog1.FileName), xls1) 'Add the original file to it
				AddHandler Work.LoadLinkedFile, AddressOf Work_LoadLinkedFile 'Set up an event to load the linked files.
																							 ' /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


				CompareXls(xls1, xls2, Nothing)
			Finally
				compareWithExcel.Enabled = True
				validateRecalc.Enabled = True
			End Try
		End Sub

		Private Sub CompareXls(ByVal xls1 As XlsFile, ByVal xls2 As XlsFile, ByVal table As DataTable)
			Dim DiffCount As Integer = 0
			xls1.Recalc()

			For sheet As Integer = 1 To xls1.SheetCount
				xls1.ActiveSheet = sheet
				xls2.ActiveSheet = sheet
				Dim aColCount As Integer = xls1.ColCount
				For r As Integer = 1 To xls1.RowCount
					For c As Integer = 1 To aColCount
						Dim f As TFormula = TryCast(xls1.GetCellValue(r, c), TFormula)
						If f IsNot Nothing Then
							Dim ad As New TCellAddress(r, c)
							Dim f2 As TFormula = CType(xls2.GetCellValue(r, c), TFormula)
							If f.Result Is Nothing Then
								f.Result = ""
							End If
							If f2.Result Is Nothing Then
								f2.Result = ""
							End If
							Dim eps As Double = 0
							If TypeOf f.Result Is Double AndAlso TypeOf f2.Result Is Double Then
								If CDbl(f2.Result) = 0 Then
									If Math.Abs(CDbl(f.Result)) < Double.Epsilon Then
										eps = 0
									Else
										eps = Double.NaN
									End If
								Else
									eps = CDbl(f.Result) / CDbl(f2.Result)
								End If
								If Math.Abs(eps - 1) < 0.001 Then
									f.Result = f2.Result
								End If
							End If
							If Not f.Result.Equals(f2.Result) Then
								If table Is Nothing Then
									report.Text &= vbLf & "Sheet:" & xls1.SheetName & " --- Cell:" & ad.CellRef & " --- Calculated: " & f.Result.ToString() & "    Excel: " & f2.Result.ToString() & "  dif: " & eps.ToString() & "   formula: " & f.Text
									Application.DoEvents()
								Else
									table.Rows.Add(New Object() { xls1.SheetName, ad.CellRef, f.Result.ToString(), f2.Result.ToString(), eps.ToString(), f.Text })
								End If
								DiffCount += 1

							End If
						End If
					Next c
				Next r
			Next sheet

			If table Is Nothing Then
				report.Text &= vbLf & "Finished Comparing."
				If DiffCount = 0 Then
					report.Text &= vbLf & "**********No differences found!**********"
				Else
					report.Text &= String.Format(vbLf & "  --->Found {0} differences", DiffCount)
				End If
			End If
		End Sub

		Private Sub ValidateXls(ByVal xls As XlsFile, ByVal table As DataTable)
			Dim Usl As TUnsupportedFormulaList = xls.RecalcAndVerify()
			For i As Integer = 0 To Usl.Count - 1
				table.Rows.Add(New Object() { Usl(i).FileName, Usl(i).Cell.CellRef, Usl(i).ErrorType.ToString(), Usl(i).FunctionName })
			Next i
		End Sub

		''' <summary>
		''' This is the method that will be called by the ASP.NET front end. It returns an array of bytes 
		''' with the report data, so the ASP.NET application can stream it to the client.
		''' </summary>
		''' <returns>The generated file as a byte array.</returns>
		Public Function WebRun(ByVal DataStream As Stream, ByVal FileName As String) As Byte()
			XlsReport.SetValue("Date", Date.Now)
			XlsReport.SetValue("FileName", FileName)
			Dim Data As New DataSet()
			Dim ValidateResult As DataTable = Data.Tables.Add("ValidateResult")
			ValidateResult.Columns.Add("FileName", GetType(String))
			ValidateResult.Columns.Add("CellRef", GetType(String))
			ValidateResult.Columns.Add("ErrorType", GetType(String))
			ValidateResult.Columns.Add("FunctionName", GetType(String))

			Dim CompareResult As DataTable = Data.Tables.Add("CompareResult")
			CompareResult.Columns.Add("SheetName", GetType(String))
			CompareResult.Columns.Add("CellRef", GetType(String))
			CompareResult.Columns.Add("CalcResult", GetType(String))
			CompareResult.Columns.Add("XlsResult", GetType(String))
			CompareResult.Columns.Add("Diff", GetType(String))
			CompareResult.Columns.Add("FormulaText", GetType(String))

			XlsReport.AddTable(Data)

			Dim xls1 As New XlsFile()
			Dim xls2 As New XlsFile()

			xls1.Open(DataStream)
			DataStream.Position = 0
			xls2.Open(DataStream)

			CompareXls(xls1, xls2, CompareResult)
			ValidateXls(xls1, ValidateResult)

			Dim DataPath As String = Path.Combine(Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), ".."), "..") & Path.DirectorySeparatorChar

			Using OutStream As New MemoryStream()
				Using InStream As New FileStream(DataPath & "ValidateReport.xls", FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
					XlsReport.Run(InStream, OutStream)
					Return OutStream.ToArray()
				End Using
			End Using
		End Function

		''' <summary>
		''' This event is used when there are linked files, to load them on demand.
		''' </summary>
		''' <param name="sender"></param>
		''' <param name="e"></param>
		Private Sub Work_LoadLinkedFile(ByVal sender As Object, ByVal e As LoadLinkedFileEventArgs)
			'IMPORTANT: DO NOT USE THIS METHOD IN PRODUCTION IF SECURITY IS IMPORTANT.
			'This method will access any file in your harddisk, as long as it is linked in the spreaadhseet, and
			'that could mean an IMPORTANT SECURITY RISK. You should limit the places where the app can search for 
			'linked files. Look at the "Recalculating Linked Files" in the PDF API Guide for more information.

			Dim FilePath As String = Path.Combine(Path.GetDirectoryName(openFileDialog1.FileName), e.FileName)

			If File.Exists(FilePath) Then 'If we find the path, just load the file.
				e.Xls = New XlsFile()
				e.Xls.Open(FilePath)
				Return
			End If

			'If we couldn't find the file, ask the user for its location.
			linkedFileDialog.FileName = FilePath
			If linkedFileDialog.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then 'if user cancels, e.Xls will be null, so no file will be used and an #errna error will show in the formulas.
				Return
			End If

			e.Xls = New XlsFile()
			e.Xls.Open(linkedFileDialog.FileName)

		End Sub

	End Class

End Namespace

