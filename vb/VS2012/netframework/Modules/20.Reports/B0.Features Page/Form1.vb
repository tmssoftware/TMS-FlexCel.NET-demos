Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports System.Data.OleDb
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Report
Imports FlexCel.Render
Imports FlexCel.Pdf
Imports System.Globalization

Imports System.Xml


Namespace FeaturesPage
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Public Sub New()
			InitializeComponent()
			'initialize the db.
			dbconnection.ConnectionString = dbconnection.ConnectionString.Replace("Features.mdb", Path.Combine(DataPath, "features.mdb"))
			ResizeToolbar(mainToolbar)
			FlexCelConfig.DpiForImages = 192 'Make the exports in hidpi.
		End Sub

		Private Sub ResizeToolbar(ByVal toolbar As ToolStrip)

			Using gr As Graphics = CreateGraphics()
				Dim xFactor As Double = gr.DpiX / 96.0
				Dim yFactor As Double = gr.DpiY / 96.0
				toolbar.ImageScalingSize = New Size(CInt(Fix(24 * xFactor)), CInt(Fix(24 * yFactor)))
				toolbar.Width = 0 'force a recalc of the buttons.
			End Using
		End Sub

		Private Sub button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles toolStripButton1.Click
			Close()
		End Sub

		Private Shared ReadOnly Property DataPath() As String
			Get
				Return Path.Combine(Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), ".."), "..") & Path.DirectorySeparatorChar
			End Get
		End Property

		Private ReadOnly Property ResultPath() As String
			Get
				Dim BasePath As String = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
				Return Path.Combine(BasePath, "Features")
			End Get
		End Property

		Private Function Export(ByVal data As DataSet) As XlsFile
			Using Report As New FlexCelReport(True)
				Report.AddTable(data)
				Report.SetUserFunction("Images", New ImagesImp())
				Dim Xls As New XlsFile(True)
				Xls.Open(Path.Combine(DataPath, "Features Page.template.xls"))

				Report.Run(Xls)
				Return Xls
			End Using

		End Function

		Private Function LoadDataSet() As DataSet
			Dim Result As New DataSet()
			featuresAdapter.Fill(Result, "Features")
			categoriesAdapter.Fill(Result, "Categories")
			hyperlinksAdapter.Fill(Result, "Hyperlinks")
			Result.Relations.Add(Result.Tables("Categories").Columns("CategoryId"), Result.Tables("Features").Columns("CategoryId"))
			Result.Relations.Add(Result.Tables("Features").Columns("FeaturesId"), Result.Tables("Hyperlinks").Columns("FeaturesId"))

			Return Result
		End Function

		Private Sub btnExportExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles toolStripButton4.Click

			Dim XlsPath As String = Path.Combine(ResultPath, "FeaturesFlexCel.xls")
			Using data As DataSet = LoadDataSet()
				Dim Xls As XlsFile = Export(data)

				Directory.CreateDirectory(ResultPath)
				Xls.Save(XlsPath)
			End Using

			If MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
				Process.Start(XlsPath)
			End If

		End Sub

		Private Sub btnExportHtml_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles toolStripButton3.Click
			Dim MainHtmlPath As String = Path.Combine(ResultPath, "featuresflexcel.htm")

			Using data As DataSet = LoadDataSet()
				Dim Xls As XlsFile = Export(data)

				Directory.CreateDirectory(ResultPath)
				Using html As New FlexCelHtmlExport(Xls, True)
					html.ImageResolution = 192
					html.ImageBackground = Color.White 'Since we are not setting html.FixIE6TransparentPngSupport, we must ensure tehre are no transparent images.
					Dim SheetSelector As New TStandardSheetSelector(TSheetSelectorPosition.Top)
					AddHandler SheetSelector.SheetSelectorEntry, AddressOf SheetSelector_SheetSelectorEntry
					SheetSelector.CssGeneral.Main &= "font-family:Verdana;font-size:10pt;"

					html.ExportAllVisibleSheetsAsTabs(ResultPath, "Features", ".htm", Nothing, Nothing, SheetSelector)

					'Rename the first tab so it is "featuresflexcel.htm";
					Dim Sheets() As String = html.GeneratedFiles.GetHtmlFiles()
					File.Delete(MainHtmlPath)
					File.Move(Sheets(0), MainHtmlPath)

				End Using
			End Using
			If MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
				Process.Start(MainHtmlPath)
			End If


		End Sub

		Private Shared Sub SheetSelector_SheetSelectorEntry(ByVal sender As Object, ByVal e As SheetSelectorEntryEventArgs)
			'We will rename the first sheet, so we need to update the links here.
			If e.ActiveSheet = 1 Then
				e.Link = "featuresflexcel.htm"
			End If
		End Sub

		Private Sub btnExportPDF_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles toolStripButton2.Click
			Dim PdfPath As String = Path.Combine(ResultPath, "FeaturesFlexCel.pdf")

			Using data As DataSet = LoadDataSet()
				Dim Xls As XlsFile = Export(data)
				Directory.CreateDirectory(ResultPath)

				Using pdf As New FlexCelPdfExport(Xls, True)
					Using pdfStream As New FileStream(PdfPath, FileMode.Create)
						pdf.BeginExport(pdfStream)
						pdf.FontMapping = TFontMapping.ReplaceAllFonts

						pdf.Properties.Subject = "A list of FlexCel.NET features"
						pdf.Properties.Author = "TMS Software"
						pdf.Properties.Title = "List of FlexCel.NET features"
						pdf.PageLayout = TPageLayout.Outlines
						pdf.ExportAllVisibleSheets(False, "Features")
						pdf.EndExport()
					End Using
				End Using
			End Using

			If MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
				Process.Start(PdfPath)
			End If

		End Sub


        Private Class ImagesImp
            Inherits TFlexCelUserFunction

			Public Sub New()
			End Sub

			Public Overrides Function Evaluate(ByVal parameters() As Object) As Object
				If parameters Is Nothing OrElse parameters.Length <> 1 Then
					Throw New ArgumentException("Bad parameter count in call to Images() user-defined function")
				End If

				Dim ImageFilename As String = Path.Combine(Path.Combine(DataPath, "images"), "Features" & Convert.ToString(parameters(0), CultureInfo.InvariantCulture) & ".png")
				If File.Exists(ImageFilename) Then
					Using fs As New FileStream(ImageFilename, FileMode.Open)
                        Dim Result(CInt(fs.Length) - 1) As Byte
						fs.Read(Result, 0, Result.Length)
						Return Result
					End Using
				End If

				Return Nothing
			End Function
		End Class

	End Class

End Namespace
