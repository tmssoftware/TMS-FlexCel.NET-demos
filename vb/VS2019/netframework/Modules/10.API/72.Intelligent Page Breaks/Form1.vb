Imports System.ComponentModel
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports System.IO
Imports System.Reflection
Imports System.Text

Imports FlexCel.Render

Namespace IntelligentPageBreaks
	''' <summary>
	''' Demo showing how to create intelligent page breaks with the API.
	''' </summary>
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Keywords As Dictionary(Of String, String) = CreateKeywords()


		Private Shared Function CreateKeywords() As Dictionary(Of String, String)
			' A very silly syntax highlighter. We don't have any context here, so for example "get" will be highlighted when it is a property or when it is not, but it is ok for this demo.
			Dim Result As New Dictionary(Of String, String)()

			Result.Add("private", Nothing)
			Result.Add("public", Nothing)
			Result.Add("protected", Nothing)
			Result.Add("internal", Nothing)
			Result.Add("static", Nothing)
			Result.Add("void", Nothing)
			Result.Add("get", Nothing)
			Result.Add("set", Nothing)
			Result.Add("return", Nothing)
			Result.Add("while", Nothing)
			Result.Add("for", Nothing)
			Result.Add("foreach", Nothing)
			Result.Add("using", Nothing)
			Result.Add("true", Nothing)
			Result.Add("false", Nothing)

			Return Result
		End Function

		Private ReadOnly Property PathToExe() As String
			Get
				Return Path.Combine(Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), ".."), "..") & Path.DirectorySeparatorChar
			End Get
		End Property

		Private Function SyntaxColor(ByVal Xls As ExcelFile, ByVal NormalFont As Integer, ByVal CommentFont As Integer, ByVal HighlightFont As Integer, ByVal line As String) As TRichString
			Dim RTFRunList As New List(Of TRTFRun)()

			Dim i As Integer = 0
			Do While i < line.Length
				If i > 0 AndAlso line.Chars(i - 1) = "/"c AndAlso line.Chars(i) = "/"c Then
					Dim rtf As TRTFRun
					rtf.FirstChar = i - 1
					rtf.FontIndex = CommentFont
					RTFRunList.Add(rtf)
					Return New TRichString(line, RTFRunList.ToArray(), Xls)

				End If

				Dim start As Integer = i
				Do While i < line.Length AndAlso Char.IsLetterOrDigit(line.Chars(i))
					i += 1
				Loop

				If i > start AndAlso Keywords.ContainsKey(line.Substring(start, i - start)) Then
					Dim rtf As TRTFRun
					rtf.FirstChar = start
					rtf.FontIndex = HighlightFont
					RTFRunList.Add(rtf)
					rtf.FirstChar = i
					rtf.FontIndex = NormalFont
					RTFRunList.Add(rtf)
				End If

				i += 1
			Loop


			Return New TRichString(line, RTFRunList.ToArray(), Xls)
		End Function

		Private Sub DumpFile(ByVal Xls As ExcelFile, ByRef Row As Integer)
			Dim fnt As TFlxFont = Xls.GetDefaultFont
			fnt.Color = Color.Blue
			Dim HighlightFont As Integer = Xls.AddFont(fnt)
			fnt.Color = Color.Green
			Dim CommentFont As Integer = Xls.AddFont(fnt)

			Dim Level As Integer = 0
			Dim LevelStart As New Stack(Of Integer)()
			LevelStart.Push(Row)

			Using sr As New StreamReader(Path.Combine(PathToExe, "Form1.cs"))
				Dim line As String
				line = sr.ReadLine()
				Do While line IsNot Nothing
					'Find the level of "keep together" for the row. We will use #region and "{" delimiters
					'to increase the level. If possible, we would want those blocks together in one page.
					Dim s As String = line.Trim()
					If s.StartsWith("#region") Then
						Level += 1
						LevelStart.Push(Row)
					End If
					If s = "{" Then
						Level += 1
						LevelStart.Push(Row - 1) 'On {} blocks, we want to keep lines together starting with the previous statement.
					End If

					If s = "#endregion" OrElse s = "}" Then
						Level -= 1
						Xls.KeepRowsTogether(LevelStart.Pop(), Row, Level + 1, False)
					End If

					Xls.KeepRowsTogether(Row, Row, Level, True)


					Xls.SetCellValue(Row, 1, SyntaxColor(Xls, 0, CommentFont, HighlightFont, line.Replace(vbTab, "    ")))
					Row += 1
					line = sr.ReadLine()
				Loop
			End Using
		End Sub

		Private Sub AddData(ByVal Xls As ExcelFile)

			'Fill the file with the contents of this c# file, many times so we can see many page breaks.
			Dim Row As Integer = 3
			DumpFile(Xls, Row)

			Xls.AutofitRowsOnWorkbook(False, True, 1)
			Xls.AutoPageBreaks(50, 100) ' we will use a 100% of page scale since we are printing to pdf.
										 'If this was to create an Excel file, pagescale should be lower to 
										 'compensate the differences between page sizes in diiferent printers in Excel

			'Export the file to PDF so we can see the page breaks.
			Using pdf As New FlexCelPdfExport(Xls, True)
				pdf.Export(saveFileDialog1.FileName)
			End Using

		End Sub


		Private Sub button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles button1.Click
			AutoRun()
		End Sub

		Public Sub AutoRun()
			If saveFileDialog1.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
				Return
			End If
			Dim Xls As ExcelFile = New XlsFile(True)
			Xls.NewFile(1, TExcelFileFormat.v2019)
			Xls.SetColWidth(1, 78 * 256) ';make longer lines wrap in the cell.
			Dim fmt As TFlxFormat = Xls.GetFormat(Xls.GetColFormat(1))
			fmt.WrapText = True

			Xls.SetColFormat(1, Xls.AddFormat(fmt))
			AddData(Xls)
			If MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
				Process.Start(saveFileDialog1.FileName)
			End If

		End Sub
	End Class
End Namespace
