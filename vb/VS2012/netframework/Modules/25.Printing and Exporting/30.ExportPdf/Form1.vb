Imports System.Collections
Imports System.ComponentModel
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports System.IO
Imports System.Reflection
Imports System.Drawing.Drawing2D
Imports FlexCel.Pdf

'only needed if you want to go unmanaged.
Imports System.Runtime.InteropServices


Namespace ExportPdf
	''' <summary>
	''' Exporting xls files to pdf.
	''' </summary>
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Private WithEvents flexCelPdfExport1 As FlexCel.Render.FlexCelPdfExport

		Public Sub New()
			InitializeComponent()
			cbFontMapping.SelectedIndex = 1
			cbPdfType.SelectedIndex = 0
			cbTagged.SelectedIndex = 0
			cbVersion.SelectedIndex = 1
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

		Private Sub button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
			Close()
		End Sub

		Private Sub LoadSheetConfig()
			Dim Xls As ExcelFile = flexCelPdfExport1.Workbook

			chGridLines.Checked = Xls.PrintGridLines
			chFormulaText.Checked = Xls.ShowFormulaText

			chPrintLeft.Checked = (Xls.PrintOptions And TPrintOptions.LeftToRight) <> 0
			edHeader.Text = Xls.PageHeader
			edFooter.Text = Xls.PageFooter
			chFitIn.Checked = Xls.PrintToFit
			edHPages.Text = Xls.PrintNumberOfHorizontalPages.ToString()
			edVPages.Text = Xls.PrintNumberOfVerticalPages.ToString()
			edVPages.ReadOnly = Not chFitIn.Checked
			edHPages.ReadOnly = Not chFitIn.Checked

			edZoom.ReadOnly = chFitIn.Checked
			edZoom.Text = Xls.PrintScale.ToString()

			Dim m As TXlsMargins = Xls.GetPrintMargins()
			edl.Text = m.Left.ToString()
			edt.Text = m.Top.ToString()
			edr.Text = m.Right.ToString()
			edb.Text = m.Bottom.ToString()
			edf.Text = m.Footer.ToString()
			edh.Text = m.Header.ToString()

			chLandscape.Checked = Xls.PrintLandscape

			edAuthor.Text = Convert.ToString(Xls.DocumentProperties.GetStandardProperty(TPropertyId.Author))
			edTitle.Text = Convert.ToString(Xls.DocumentProperties.GetStandardProperty(TPropertyId.Title))
			edSubject.Text = Convert.ToString(Xls.DocumentProperties.GetStandardProperty(TPropertyId.Subject))
		End Sub

		Private Sub openFile_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles openFile.Click
			If openFileDialog1.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
				Return
			End If
			flexCelPdfExport1.Workbook = New XlsFile()

			flexCelPdfExport1.Workbook.Open(openFileDialog1.FileName)

			edFileName.Text = openFileDialog1.FileName

			Dim Xls As ExcelFile = flexCelPdfExport1.Workbook

			cbSheet.Items.Clear()
			Dim ActSheet As Integer = Xls.ActiveSheet
			For i As Integer = 1 To Xls.SheetCount
				Xls.ActiveSheet = i
				cbSheet.Items.Add(Xls.SheetName)
			Next i
			Xls.ActiveSheet = ActSheet
			cbSheet.SelectedIndex = ActSheet - 1

			LoadSheetConfig()
		End Sub

		Private Function HasFileOpen() As Boolean
			If flexCelPdfExport1.Workbook Is Nothing Then
				MessageBox.Show("You need to open a file first.")
				Return False
			End If
			Return True
		End Function

		Private Function LoadPreferences() As Boolean
			'NOTE: THERE SHOULD BE *A LOT* MORE VALIDATION OF VALUES ON THIS METHOD. (For example, validate that margins are between bounds)
			' As this is a simple demo, they are not included. 
			Try
				Dim Xls As ExcelFile = flexCelPdfExport1.Workbook
				Xls.PrintGridLines = chGridLines.Checked
				Xls.PageHeader = edHeader.Text
				Xls.PageFooter = edFooter.Text
				Xls.ShowFormulaText = chFormulaText.Checked

				If chFitIn.Checked Then
					Xls.PrintToFit = True
					Xls.PrintNumberOfHorizontalPages = Convert.ToInt32(edHPages.Text)
					Xls.PrintNumberOfVerticalPages = Convert.ToInt32(edVPages.Text)
				Else
					Xls.PrintToFit = False
				End If

				If chPrintLeft.Checked Then
					Xls.PrintOptions = Xls.PrintOptions Or TPrintOptions.LeftToRight
				Else
					Xls.PrintOptions = Xls.PrintOptions And Not TPrintOptions.LeftToRight
				End If

				Try
					Xls.PrintScale = Convert.ToInt32(edZoom.Text)
				Catch
					MessageBox.Show("Invalid Zoom")
					Return False
				End Try

				Dim m As New TXlsMargins()
				m.Left = Convert.ToDouble(edl.Text)
				m.Top = Convert.ToDouble(edt.Text)
				m.Right = Convert.ToDouble(edr.Text)
				m.Bottom = Convert.ToDouble(edb.Text)
				m.Footer = Convert.ToDouble(edf.Text)
				m.Header = Convert.ToDouble(edh.Text)
				Xls.SetPrintMargins(m)

				flexCelPdfExport1.PrintRangeLeft = Convert.ToInt32(edLeft.Text)
				flexCelPdfExport1.PrintRangeTop = Convert.ToInt32(edTop.Text)
				flexCelPdfExport1.PrintRangeRight = Convert.ToInt32(edRight.Text)
				flexCelPdfExport1.PrintRangeBottom = Convert.ToInt32(edBottom.Text)

				If chEmbed.Checked Then
					flexCelPdfExport1.FontEmbed = TFontEmbed.Embed
				Else
					flexCelPdfExport1.FontEmbed = TFontEmbed.None
				End If

				If chSubset.Checked Then
					flexCelPdfExport1.FontSubset = TFontSubset.Subset
				Else
					flexCelPdfExport1.FontSubset = TFontSubset.DontSubset
				End If

				flexCelPdfExport1.Kerning = cbKerning.Checked

				Select Case cbFontMapping.SelectedIndex
					Case 0
						flexCelPdfExport1.FontMapping = TFontMapping.ReplaceAllFonts
					Case 1
						flexCelPdfExport1.FontMapping = TFontMapping.ReplaceStandardFonts
					Case 2
						flexCelPdfExport1.FontMapping = TFontMapping.DontReplaceFonts
				End Select

				Select Case cbPdfType.SelectedIndex
					Case 0
						flexCelPdfExport1.PdfType = TPdfType.Standard
					Case 1
						flexCelPdfExport1.PdfType = TPdfType.PDFA1
					Case 2
						flexCelPdfExport1.PdfType = TPdfType.PDFA2
					Case 3
						flexCelPdfExport1.PdfType = TPdfType.PDFA3
				End Select

				Select Case cbTagged.SelectedIndex
					Case 0
						flexCelPdfExport1.TagMode = TTagMode.Full
					Case 1
						flexCelPdfExport1.TagMode = TTagMode.None
				End Select

				Select Case cbVersion.SelectedIndex
					Case 0
						flexCelPdfExport1.PdfVersion = TPdfVersion.v14
					Case 1
						flexCelPdfExport1.PdfVersion = TPdfVersion.v16
				End Select

				flexCelPdfExport1.Properties.Author = edAuthor.Text
				flexCelPdfExport1.Properties.Title = edTitle.Text
				flexCelPdfExport1.Properties.Subject = edSubject.Text
				flexCelPdfExport1.Properties.Language = edLang.Text

				Xls.PrintLandscape = chLandscape.Checked
			Catch e As Exception
				MessageBox.Show("Error: " & e.Message)
				Return False
			End Try
			Return True
		End Function


		Private Sub chFitIn_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chFitIn.CheckedChanged
			edVPages.ReadOnly = Not chFitIn.Checked
			edHPages.ReadOnly = Not chFitIn.Checked
			edZoom.ReadOnly = chFitIn.Checked
		End Sub

		Private Sub cbSheet_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbSheet.SelectedIndexChanged
			flexCelPdfExport1.Workbook.ActiveSheet = cbSheet.SelectedIndex + 1
			LoadSheetConfig()
		End Sub

		Private Sub export_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles export.Click
			If Not HasFileOpen() Then
				Return
			End If
			If Not LoadPreferences() Then
				Return
			End If

			If exportDialog.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
				Return
			End If

			Using Pdf As New FileStream(exportDialog.FileName, FileMode.Create)
				Dim SaveSheet As Integer = flexCelPdfExport1.Workbook.ActiveSheet
				Try
					flexCelPdfExport1.BeginExport(Pdf)
					If chExportAll.Checked Then
						flexCelPdfExport1.PageLayout = TPageLayout.Outlines 'To how the bookmarks when opening the file.
						flexCelPdfExport1.ExportAllVisibleSheets(cbResetPageNumber.Checked, Path.GetFileNameWithoutExtension(exportDialog.FileName))
					Else
						flexCelPdfExport1.PageLayout = TPageLayout.None
						flexCelPdfExport1.ExportSheet()
					End If
					flexCelPdfExport1.EndExport()
				Finally
					flexCelPdfExport1.Workbook.ActiveSheet = SaveSheet
				End Try

				If MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
					Process.Start(exportDialog.FileName)
				End If
			End Using
		End Sub

		Private Sub chExportAll_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chExportAll.CheckedChanged
			cbSheet.Enabled = Not chExportAll.Checked
			cbResetPageNumber.Enabled = chExportAll.Checked
		End Sub

		''' <summary>
		''' Add a "Confidential" watermark on each page.
		''' </summary>
		''' <param name="sender"></param>
		''' <param name="e"></param>
		Private Sub flexCelPdfExport1_AfterGeneratePage(ByVal sender As Object, ByVal e As FlexCel.Render.PageEventArgs) Handles flexCelPdfExport1.AfterGeneratePage
			If Not cbConfidential.Checked Then
				Return
			End If

			Const s As String = "Confidential"

			Using ABrush As Brush = New SolidBrush(Color.FromArgb(30, 25, 25, 25)) 'Red=Green=Blue is a shade of gray. Alpha=30 means it is transparent (255 is pure opaque, 0 is pure transparent).
                Using AFont As TUIFont = TUIFont.Create("Arial", 72)
                    Dim x0 As Double = e.File.PageSize.Width * 72.0 / 100.0 / 2.0 'PageSize is in inches/100, our coordinate system is in Points, that is inches/72
                    Dim y0 As Double = e.File.PageSize.Height * 72.0 / 100.0 / 2.0
                    Dim sf As SizeF = e.File.MeasureString(s, AFont)
                    e.File.Rotate(x0, y0, 45)
                    e.File.DrawString(s, AFont, ABrush, x0 - sf.Width / 2.0, y0 + sf.Height / 2.0) 'the y coord means the bottom of the text, and as the y axis grows down, we have to add sf.height/2 instead of substracting it.
                End Using
            End Using
		End Sub


		''' <summary>
		''' We show on this event how you can make an unmanaged call to the Win32 API to return font information and avoid
		''' scanning the "fonts" folder. Note that this is <b>UNMANAGED</b> code, and it is not really needed except for small performance concerns,
		''' so avoid using it if you don't really need it. Please read UsingFlexCelPdfExport for more information.
		''' </summary>
		''' <param name="sender"></param>
		''' <param name="e"></param>
		Private Sub flexCelPdfExport1_GetFontData(ByVal sender As Object, ByVal e As FlexCel.Pdf.GetFontDataEventArgs) Handles flexCelPdfExport1.GetFontData
			'If the checkbox is not checked, just ignore this event.
			If Not cbUseGetFontData.Checked Then
				e.Applied = False
				Return
			End If

			'Actually make the WIN32 call.
			Dim ttcf As UInteger = &H66637474 'return full true type collections.

			' Allocate a handle for the font
			Dim FontHandle As IntPtr = CType(e.InputFont, FlexCel.Draw.TGdipUIFont).Handle.ToHfont()
			Try
				Using Gr As Graphics = Graphics.FromHwnd(IntPtr.Zero)
					Dim GrHandle As IntPtr = Gr.GetHdc()
					Try
						Dim ObjHandle As IntPtr = SelectObject(GrHandle, FontHandle)
						Try
							'First find out the sizes
							Dim Size As UInteger = GetFontData(GrHandle, ttcf, 0, Nothing, 0)
							If CInt(Size) < 0 Then 'error
								ttcf = 0 'This might not be a true type collection, try again.
								Size = GetFontData(GrHandle, ttcf, 0, Nothing, 0)

								If CInt(Size) < 0 Then 'nothing else to do, exit.
									e.Applied = False
									Return
								End If
							End If

							'Now get the font data.
							e.FontData = New Byte(CInt(Size) - 1){}
							Dim Result As UInteger = GetFontData(GrHandle, ttcf, 0, e.FontData, Size)

							If CInt(Result) < 0 Then
								e.Applied = False
								Return
							End If
							e.Applied = True
						Finally
							DeleteObject(ObjHandle)
						End Try
					Finally
						Gr.ReleaseHdc(GrHandle)
					End Try
				End Using
			Finally
				DeleteObject(FontHandle)
			End Try
		End Sub

		''' <summary>
		''' The Win32 call.
		''' </summary>
		''' <param name="hdc"></param>
		''' <param name="dwTable"></param>
		''' <param name="dwOffset"></param>
		''' <param name="lpvBuffer"></param>
		''' <param name="cbData"></param>
		''' <returns></returns>
		<DllImport("gdi32.dll")> _
		Shared Function GetFontData(ByVal hdc As IntPtr, ByVal dwTable As UInteger, ByVal dwOffset As UInteger, <[In](), Out()> ByVal lpvBuffer() As Byte, ByVal cbData As UInteger) As UInteger
		End Function

		<DllImport("GDI32.dll")> _
		Shared Function DeleteObject(ByVal objectHandle As IntPtr) As Boolean
		End Function

		<DllImport("gdi32.dll")> _
		Shared Function SelectObject(ByVal hdc As IntPtr, ByVal hgdiobj As IntPtr) As IntPtr
		End Function




	End Class
End Namespace
