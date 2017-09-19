Imports System.Collections
Imports System.ComponentModel
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Render
Imports System.IO

Imports System.Text



Namespace ExportHTML
	''' <summary>
	''' An Example on how to export to HTML.
	''' </summary>
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Private WithEvents flexCelHtmlExport1 As FlexCel.Render.FlexCelHtmlExport


		Public Sub New()
			InitializeComponent()
		End Sub

		Private MailDialog As Mailform

		Private Sub button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
			Close()
		End Sub

		Private Sub LoadSheetConfig()
			Dim Xls As ExcelFile = flexCelHtmlExport1.Workbook

			chGridLines.Checked = Xls.PrintGridLines
			chPrintHeadings.Checked = Xls.PrintHeadings
			chFormulaText.Checked = Xls.ShowFormulaText
		End Sub

		Private Sub openFile_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles openFile.Click
			If openFileDialog1.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
				Return
			End If
			flexCelHtmlExport1.Workbook = New XlsFile()

			flexCelHtmlExport1.Workbook.Open(openFileDialog1.FileName)

			Text = "Export: " & openFileDialog1.FileName

			Dim Xls As ExcelFile = flexCelHtmlExport1.Workbook

			cbSheet.Items.Clear()
			For i As Integer = 1 To Xls.SheetCount
				cbSheet.Items.Add(Xls.GetSheetName(i))
			Next i
			cbSheet.SelectedIndex = Xls.ActiveSheet - 1

			LoadSheetConfig()
		End Sub

		Private Function HasFileOpen() As Boolean
			If flexCelHtmlExport1.Workbook Is Nothing Then
				MessageBox.Show("You need to open a file first.")
				Return False
			End If
			Return True
		End Function

		Private Function LoadPreferences() As Boolean
			'NOTE: THERE SHOULD BE *A LOT* MORE VALIDATION OF VALUES ON THIS METHOD. (For example, validate that margins are between bounds)
			' As this is a simple demo, they are not included. 
			Try
				Dim Xls As ExcelFile = flexCelHtmlExport1.Workbook

				'Note: In this demo we will only apply this things to the active sheet.
				'If you want to apply the settings to all the sheets, you should loop in the sheets and change them here.
				Xls.PrintGridLines = chGridLines.Checked
				Xls.PrintHeadings = chPrintHeadings.Checked
				Xls.ShowFormulaText = chFormulaText.Checked

				flexCelHtmlExport1.PrintRangeLeft = Convert.ToInt32(edLeft.Text)
				flexCelHtmlExport1.PrintRangeTop = Convert.ToInt32(edTop.Text)
				flexCelHtmlExport1.PrintRangeRight = Convert.ToInt32(edRight.Text)
				flexCelHtmlExport1.PrintRangeBottom = Convert.ToInt32(edBottom.Text)

				If sbSVG.Checked Then
					flexCelHtmlExport1.SavedImagesFormat = THtmlImageFormat.Svg
				Else
					flexCelHtmlExport1.SavedImagesFormat = THtmlImageFormat.Png
				End If
				flexCelHtmlExport1.EmbedImages = cbEmbedImages.Checked

				flexCelHtmlExport1.FixOutlook2007CssSupport = cbOutlook2007.Checked
				flexCelHtmlExport1.FixIE6TransparentPngSupport = cbIe6Png.Checked

				flexCelHtmlExport1.HidePrintObjects = THidePrintObjects.None
				If Not cbImages.Checked Then
					flexCelHtmlExport1.HidePrintObjects = flexCelHtmlExport1.HidePrintObjects Or THidePrintObjects.Images
				End If
				If Not cbHyperlinks.Checked Then
					flexCelHtmlExport1.HidePrintObjects = flexCelHtmlExport1.HidePrintObjects Or THidePrintObjects.Hyperlynks
				End If
				If Not cbComments.Checked Then
					flexCelHtmlExport1.HidePrintObjects = flexCelHtmlExport1.HidePrintObjects Or THidePrintObjects.Comments
				End If
				If Not cbHeadersFooters.Checked Then
					flexCelHtmlExport1.HidePrintObjects = flexCelHtmlExport1.HidePrintObjects Or THidePrintObjects.HeadersAndFooters
				End If

			Catch e As Exception
				MessageBox.Show("Error: " & e.Message)
				Return False
			End Try
			Return True
		End Function


		Private Sub cbSheet_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
			flexCelHtmlExport1.Workbook.ActiveSheet = cbSheet.SelectedIndex + 1
			LoadSheetConfig()
		End Sub

		Private Sub export_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles export.Click
			If Not HasFileOpen() Then
				Return
			End If
			If Not LoadPreferences() Then
				Return
			End If

			If cbFileFormat.SelectedIndex = 1 Then
				flexCelHtmlExport1.HtmlFileFormat = THtmlFileFormat.MHtml
				exportDialog.FilterIndex = 2
			Else
				flexCelHtmlExport1.HtmlFileFormat = THtmlFileFormat.Html
				exportDialog.FilterIndex = 1
			End If

			If exportDialog.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
				Return
			End If

			flexCelHtmlExport1.AllowOverwritingFiles = True

			Dim CssFileName As String = Nothing
			If cbCss.Checked Then
				CssFileName = edCss.Text
			End If

			Dim FileNameToOpen As String = exportDialog.FileName

			Select Case cbHtmlVersion.SelectedIndex
				Case 0
					flexCelHtmlExport1.HtmlVersion = THtmlVersion.Html_32
				Case 2
					flexCelHtmlExport1.HtmlVersion = THtmlVersion.XHTML_10
				Case 3
					flexCelHtmlExport1.HtmlVersion = THtmlVersion.Html_5
				Case Else
					flexCelHtmlExport1.HtmlVersion = THtmlVersion.Html_401
			End Select

			If edBodyStart.Text IsNot Nothing Then
				flexCelHtmlExport1.ExtraInfo.BodyStart = New String() { edBodyStart.Text }
			End If

			Select Case cbExportObject.SelectedIndex
				Case 0
					Dim SelectorPosition As TSheetSelectorPosition = TSheetSelectorPosition.None

					'If in VB.NET or Delphi.NET, use "if cbTop.Checked then SelectorPosition = SelectorPosition or TSheetSelectorPosition.Top"
					If cbTop.Checked Then
						SelectorPosition = SelectorPosition Or TSheetSelectorPosition.Top
					End If
					If cbLeft.Checked Then
						SelectorPosition = SelectorPosition Or TSheetSelectorPosition.Left
					End If
					If cbBottom.Checked Then
						SelectorPosition = SelectorPosition Or TSheetSelectorPosition.Bottom
					End If
					If cbRight.Checked Then
						SelectorPosition = SelectorPosition Or TSheetSelectorPosition.Right
					End If


					flexCelHtmlExport1.ExportAllVisibleSheetsAsTabs(Path.GetDirectoryName(exportDialog.FileName), Path.GetFileNameWithoutExtension(exportDialog.FileName), Path.GetExtension(exportDialog.FileName), edImages.Text, CssFileName, New TStandardSheetSelector(SelectorPosition))

					FileNameToOpen = Path.Combine(Path.GetDirectoryName(exportDialog.FileName), Path.GetFileNameWithoutExtension(exportDialog.FileName))
					FileNameToOpen = Path.Combine(FileNameToOpen, flexCelHtmlExport1.Workbook.SheetName)
					FileNameToOpen = Path.Combine(FileNameToOpen, Path.GetExtension(exportDialog.FileName))

				Case 1
					flexCelHtmlExport1.ExportAllVisibleSheetsAsOneHtmlFile(exportDialog.FileName, edImages.Text, CssFileName, edSheetSeparator.Text)

				Case 2
						flexCelHtmlExport1.Export(exportDialog.FileName, edImages.Text, CssFileName)
						Exit Select
			End Select

			Dim GeneratedFiles() As String = flexCelHtmlExport1.GeneratedFiles.GetHtmlFiles()
			If GeneratedFiles.Length = 0 Then
				MessageBox.Show("Error: No file has been generated")
			Else
				If MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
					Process.Start(GeneratedFiles(0))
				End If
			End If
		End Sub

		Private Sub btnEmail_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEmail.Click
			If Not HasFileOpen() Then
				Return
			End If
			If MailDialog Is Nothing Then
				MailDialog = New Mailform()
			End If
			MailDialog.MainForm = Me

			If Not flexCelHtmlExport1.FixOutlook2007CssSupport Then
				Dim dr As DialogResult = MessageBox.Show("You have not checked ""Outlook 2007 support"". If any of your clients has Outlook express, you should turn this on." & vbLf & vbLf & "Use Outlook 2007 fix?", "Warning", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning)

				If dr = System.Windows.Forms.DialogResult.Cancel Then
					Return
				End If
				If dr = System.Windows.Forms.DialogResult.Yes Then
					cbOutlook2007.Checked = True
					flexCelHtmlExport1.FixOutlook2007CssSupport = True
				End If
			End If

			MailDialog.ShowDialog()

		End Sub

		Public Function GenerateMHTML() As Byte()
			LoadPreferences()
			flexCelHtmlExport1.HtmlFileFormat = THtmlFileFormat.MHtml


			flexCelHtmlExport1.AllowOverwritingFiles = True

			flexCelHtmlExport1.HtmlVersion = THtmlVersion.Html_401

			If edBodyStart.Text IsNot Nothing Then
				flexCelHtmlExport1.ExtraInfo.BodyStart = New String() { edBodyStart.Text }
			End If

			Using ms As New MemoryStream()
				Using writer As New StreamWriter(ms, Encoding.UTF8)
					flexCelHtmlExport1.Export(writer, flexCelHtmlExport1.Workbook.ActiveFileName, Nothing)
				End Using
				Return ms.ToArray()
			End Using
		End Function


		Private Sub mainForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
			cbExportObject.SelectedIndex = 0
			cbHtmlVersion.SelectedIndex = 3
			cbFileFormat.SelectedIndex = 0
		End Sub

		Private Sub cbExportObject_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbExportObject.SelectedIndexChanged
			edSheetSeparator.Enabled = cbExportObject.SelectedIndex = 1
			cbTop.Enabled = cbExportObject.SelectedIndex = 0
			cbLeft.Enabled = cbExportObject.SelectedIndex = 0
			cbRight.Enabled = cbExportObject.SelectedIndex = 0
			cbBottom.Enabled = cbExportObject.SelectedIndex = 0
			cbSheet.Enabled = cbExportObject.SelectedIndex = 2
		End Sub

		Private Sub cbCss_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbCss.CheckedChanged
			edCss.Enabled = cbCss.Checked
		End Sub

		Private Sub flexCelHtmlExport1_HtmlFont(ByVal sender As Object, ByVal e As FlexCel.Core.HtmlFontEventArgs) Handles flexCelHtmlExport1.HtmlFont
			If cbReplaceFonts.Checked Then
				e.FontFamily = "arial, sans-serif"
			End If
		End Sub

	End Class
End Namespace
