Imports System.Drawing.Imaging
Imports System.Collections
Imports System.ComponentModel

Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Render

Imports System.IO
Imports System.Reflection
Imports System.Drawing.Drawing2D
Imports System.Drawing.Printing

Imports System.Runtime.InteropServices
Imports FlexCel.Draw

Namespace PrintPreviewandExport
	''' <summary>
	''' Printing / Previewing and Exporting xls files.
	''' </summary>
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Private WithEvents flexCelPrintDocument1 As FlexCel.Render.FlexCelPrintDocument

		Public Sub New()
			InitializeComponent()
			cbInterpolation.SelectedIndex = 1
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

		Private Sub LoadSheetConfig()
			Dim Xls As ExcelFile = flexCelPrintDocument1.Workbook

			chGridLines.Checked = Xls.PrintGridLines
			chHeadings.Checked = Xls.PrintHeadings
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

			Landscape.Checked = (Xls.PrintOptions And TPrintOptions.Orientation) = 0

		End Sub

		Private Sub openFile_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOpenFile.Click
			If openFileDialog1.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
				Return
			End If
			flexCelPrintDocument1.Workbook = New XlsFile()

			flexCelPrintDocument1.Workbook.Open(openFileDialog1.FileName)

			edFileName.Text = openFileDialog1.FileName

			Dim Xls As ExcelFile = flexCelPrintDocument1.Workbook

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
			If flexCelPrintDocument1.Workbook Is Nothing Then
				MessageBox.Show("You need to open a file first.")
				Return False
			End If
			Return True
		End Function

		Private Function LoadPreferences() As Boolean
			'NOTE: THERE SHOULD BE *A LOT* MORE VALIDATION OF VALUES ON THIS METHOD. (For example, validate that margins are between bounds)
			' As this is a simple demo, they are not included. 
			Try
				flexCelPrintDocument1.AllVisibleSheets = cbAllSheets.Checked
				flexCelPrintDocument1.ResetPageNumberOnEachSheet = cbResetPageNumber.Checked
				flexCelPrintDocument1.AntiAliasedText = chAntiAlias.Checked

				Dim Xls As ExcelFile = flexCelPrintDocument1.Workbook
				Xls.PrintGridLines = chGridLines.Checked
				Xls.PrintHeadings = chHeadings.Checked
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


				flexCelPrintDocument1.PrintRangeLeft = Convert.ToInt32(edLeft.Text)
				flexCelPrintDocument1.PrintRangeTop = Convert.ToInt32(edTop.Text)
				flexCelPrintDocument1.PrintRangeRight = Convert.ToInt32(edRight.Text)
				flexCelPrintDocument1.PrintRangeBottom = Convert.ToInt32(edBottom.Text)

				flexCelPrintDocument1.DocumentName = flexCelPrintDocument1.Workbook.ActiveFileName & " - Sheet " & flexCelPrintDocument1.Workbook.ActiveSheetByName

				flexCelPrintDocument1.DefaultPageSettings.Landscape = Landscape.Checked
			Catch e As Exception
				MessageBox.Show("Error: " & e.Message)
				Return False
			End Try
			Return True
		End Function

		Private Sub preview_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPreview.Click
			If Not HasFileOpen() Then
				Return
			End If
			If Not LoadPreferences() Then
				Return
			End If
			If Not DoSetup() Then
				Return
			End If

			'If you want to bypass the paper size selected on the dialog and use the one on Excel, uncomment
			'the following lines:
			'TPaperDimensions t= flexCelPrintDocument1.Workbook.PrintPaperDimensions;
			'flexCelPrintDocument1.DefaultPageSettings.PaperSize = new PaperSize(t.PaperName, Convert.ToInt32(t.Width), Convert.ToInt32(t.Height));

			printPreviewDialog1.ShowDialog()

		End Sub

		Private Function DoSetup() As Boolean
			Dim Result As Boolean = printDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK
			Landscape.Checked = flexCelPrintDocument1.DefaultPageSettings.Landscape
			Return Result
		End Function

		Private Sub setup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSetup.Click
			DoSetup()
		End Sub

		Private Sub chFitIn_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chFitIn.CheckedChanged
			edVPages.ReadOnly = Not chFitIn.Checked
			edHPages.ReadOnly = Not chFitIn.Checked
			edZoom.ReadOnly = chFitIn.Checked
		End Sub

		Private Sub print_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrint.Click
			If Not HasFileOpen() Then
				Return
			End If
			If Not LoadPreferences() Then
				Return
			End If
			If Not DoSetup() Then
				Return
			End If
			flexCelPrintDocument1.Print()
		End Sub

		Private Sub cbSheet_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbSheet.SelectedIndexChanged
			flexCelPrintDocument1.Workbook.ActiveSheet = cbSheet.SelectedIndex + 1
			LoadSheetConfig()

		End Sub


		''' <summary>
		''' Add a "Confidential" watermark on each page.
		''' </summary>
		''' <param name="sender"></param>
		''' <param name="e"></param>
		Private Sub flexCelPrintDocument1_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles flexCelPrintDocument1.PrintPage
			If Not cbConfidential.Checked Then
				Return
			End If

			Using myMatrix As New Matrix()
				myMatrix.RotateAt(45, New PointF(e.PageBounds.Left + e.MarginBounds.Width / 2F, e.PageBounds.Top + e.MarginBounds.Height / 2F), MatrixOrder.Append)
				e.Graphics.Transform = myMatrix
			End Using

			Using ABrush As Brush = New SolidBrush(Color.FromArgb(30, 25, 25, 25)) 'Red=Green=Blue is a shade of gray. Alpha=30 means it is transparent (255 is pure opaque, 0 is pure transparent).
				Using AFont As New Font("Arial", 72)
					Using sf As New StringFormat()

						sf.Alignment = StringAlignment.Center
						sf.LineAlignment = StringAlignment.Center
						e.Graphics.DrawString("Confidential", AFont, ABrush, e.PageBounds, sf)
					End Using
				End Using
			End Using
		End Sub

		#Region "Hard Margins"
		'Shows how to read the hard margins from a printer if you really need to.

		Private Sub flexCelPrintDocument1_BeforePrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles flexCelPrintDocument1.BeforePrintPage
			Select Case cbInterpolation.SelectedIndex
				Case 0
					e.Graphics.InterpolationMode = InterpolationMode.Bicubic
				Case 1
					e.Graphics.InterpolationMode = InterpolationMode.Bilinear
				Case 2
					e.Graphics.InterpolationMode = InterpolationMode.Default
				Case 3
					e.Graphics.InterpolationMode = InterpolationMode.High
				Case 4
					e.Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic
				Case 5
					e.Graphics.InterpolationMode = InterpolationMode.HighQualityBilinear
				Case 6
					e.Graphics.InterpolationMode = InterpolationMode.Low
				Case 7
					e.Graphics.InterpolationMode = InterpolationMode.NearestNeighbor
			End Select

		End Sub

		<DllImport("gdi32.dll")> _
		Private Shared Function GetDeviceCaps(ByVal hdc As IntPtr, ByVal capindex As Int32) As Int32
		End Function

		''' <summary>
		''' This event will adjust for a better position on the page for some printers. 
		''' It is not normally necessary, and it has to make an unmanaged call to GetDeviceCaps,
		''' but it is given here as an example of how it could be done.
		''' </summary>
		''' <param name="sender"></param>
		''' <param name="e"></param>
		Private Sub flexCelPrintDocument1_GetPrinterHardMargins(ByVal sender As Object, ByVal e As FlexCel.Render.PrintHardMarginsEventArgs) Handles flexCelPrintDocument1.GetPrinterHardMargins
			Const PHYSICALOFFSETX As Integer = 112
			Const PHYSICALOFFSETY As Integer = 113

			Dim DpiX As Double = e.Graphics.DpiX
			Dim DpiY As Double = e.Graphics.DpiY

			Dim Hdc As IntPtr = e.Graphics.GetHdc()
			Try
				e.XMargin = CSng(GetDeviceCaps(Hdc, PHYSICALOFFSETX) * 100.0 / DpiX)
				e.YMargin = CSng(GetDeviceCaps(Hdc, PHYSICALOFFSETY) * 100.0 / DpiY)

			Finally
				e.Graphics.ReleaseHdc(Hdc)
			End Try

		End Sub
		#End Region

		#Region "Export as image"

		#Region "Common methods to Export with FlexCelImgExport"
		Private Function CreateBitmap(ByVal Resolution As Double, ByVal pd As TPaperDimensions, ByVal PxFormat As PixelFormat) As Bitmap
			Dim Result As New Bitmap(CInt(Fix(Math.Ceiling(pd.Width / 96F * Resolution))), CInt(Fix(Math.Ceiling(pd.Height / 96F * Resolution))), PxFormat)
			Result.SetResolution(CSng(Resolution), CSng(Resolution))
			Return Result

		End Function
		#End Region

		#Region "Export using FlexCelImgExport - simple images the hard way. DO NOT USE IF NOT DESPERATE!"
		'The methods shows how to use FlexCelImgExport the "hard way", without using SaveAsImage.
		'For normal operation you should only need to call SaveAsImage, but you could use the code here
		'if you need to customize the ImgExport output, or if you need to get all the images as different files.
		Private Sub CreateImg(ByVal OutStream As Stream, ByVal ImgExport As FlexCelImgExport, ByVal ImgFormat As ImageFormat, ByVal Colors As ImageColorDepth, ByRef ExportInfo As TImgExportInfo)
			Dim pd As TPaperDimensions = ImgExport.GetRealPageSize()

			Dim RgbPixFormat As PixelFormat
			If Colors <> ImageColorDepth.TrueColor Then
				RgbPixFormat = PixelFormat.Format32bppPArgb
			Else
				RgbPixFormat = PixelFormat.Format24bppRgb
			End If
			Dim PixFormat As PixelFormat = PixelFormat.Format1bppIndexed
			Select Case Colors
				Case ImageColorDepth.TrueColor
					PixFormat = RgbPixFormat
				Case ImageColorDepth.Color256
					PixFormat = PixelFormat.Format8bppIndexed
			End Select

			Using OutImg As Bitmap = CreateBitmap(ImgExport.Resolution, pd, PixFormat)
				Dim ActualOutImg As Bitmap
				If Colors <> ImageColorDepth.TrueColor Then
					ActualOutImg = CreateBitmap(ImgExport.Resolution, pd, RgbPixFormat)
				Else
					ActualOutImg = OutImg
				End If
				Try
					Using Gr As Graphics = Graphics.FromImage(ActualOutImg)
						Gr.FillRectangle(Brushes.White, 0, 0, ActualOutImg.Width, ActualOutImg.Height) 'Clear the background
						ImgExport.ExportNext(Gr, ExportInfo)
					End Using

					If Colors = ImageColorDepth.BlackAndWhite Then
						FloydSteinbergDither.ConvertToBlackAndWhite(ActualOutImg, OutImg)
					Else
						If Colors = ImageColorDepth.Color256 Then
						OctreeQuantizer.ConvertTo256Colors(ActualOutImg, OutImg)
						End If
					End If
				Finally
					If ActualOutImg IsNot OutImg Then
						ActualOutImg.Dispose()
					End If
				End Try

				OutImg.Save(OutStream, ImgFormat)
			End Using
		End Sub

		Private Sub ExportAllImages(ByVal ImgExport As FlexCelImgExport, ByVal ImgFormat As ImageFormat, ByVal ColorDepth As ImageColorDepth)
			Dim ExportInfo As TImgExportInfo = Nothing 'For first page.
			Dim i As Integer = 0
			Do
				Dim FileName As String = Path.GetDirectoryName(exportImageDialog.FileName) & Path.DirectorySeparatorChar & Path.GetFileNameWithoutExtension(exportImageDialog.FileName) & "_" & ImgExport.Workbook.SheetName & String.Format("_{0:0000}", i) & Path.GetExtension(exportImageDialog.FileName)
				Using ImageStream As New FileStream(FileName, FileMode.Create)
					CreateImg(ImageStream, ImgExport, ImgFormat, ColorDepth, ExportInfo)
				End Using
				i += 1
			Loop While ExportInfo.CurrentPage < ExportInfo.TotalPages
		End Sub

		Private Sub DoExportUsingFlexCelImgExportComplex(ByVal ColorDepth As ImageColorDepth)
			If Not HasFileOpen() Then
				Return
			End If
			If Not LoadPreferences() Then
				Return
			End If

			If exportImageDialog.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
				Return
			End If

			Dim ImgFormat As System.Drawing.Imaging.ImageFormat = System.Drawing.Imaging.ImageFormat.Png
			If String.Compare(Path.GetExtension(exportImageDialog.FileName), ".jpg", True) = 0 Then
				ImgFormat = System.Drawing.Imaging.ImageFormat.Jpeg
			End If

			Using ImgExport As New FlexCelImgExport(flexCelPrintDocument1.Workbook)
				ImgExport.Resolution = 96 'To get a better quality image but with larger file size too, increate this value. (for example to 300 or 600 dpi)

				If cbAllSheets.Checked Then
					Dim SaveActiveSheet As Integer = ImgExport.Workbook.ActiveSheet
					Try
						ImgExport.Workbook.ActiveSheet = 1
						Dim Finished As Boolean = False
						Do While Not Finished
							ExportAllImages(ImgExport, ImgFormat, ColorDepth)
							If ImgExport.Workbook.ActiveSheet < ImgExport.Workbook.SheetCount Then
								ImgExport.Workbook.ActiveSheet += 1
							Else
								Finished = True
							End If

						Loop
					Finally
						ImgExport.Workbook.ActiveSheet = SaveActiveSheet
					End Try
				Else
					ExportAllImages(ImgExport, ImgFormat, ColorDepth)
				End If

			End Using

		End Sub
		#End Region

		#Region "Export using FlexCelImgExport - simple images the simple way."

		Private Sub DoExportUsingFlexCelImgExportSimple(ByVal ColorDepth As ImageColorDepth)
			If Not HasFileOpen() Then
				Return
			End If
			If Not LoadPreferences() Then
				Return
			End If

			If exportImageDialog.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
				Return
			End If

			Dim ImgFormat As ImageExportType = ImageExportType.Png
			If String.Compare(Path.GetExtension(exportImageDialog.FileName), ".jpg", True) = 0 Then
				ImgFormat = ImageExportType.Jpeg
			End If

			Using ImgExport As New FlexCelImgExport(flexCelPrintDocument1.Workbook)
				ImgExport.AllVisibleSheets = cbAllSheets.Checked
				ImgExport.ResetPageNumberOnEachSheet = cbResetPageNumber.Checked
				ImgExport.Resolution = 96 'To get a better quality image but with larger file size too, increate this value. (for example to 300 or 600 dpi)
				ImgExport.SaveAsImage(exportImageDialog.FileName, ImgFormat, ColorDepth)
			End Using
		End Sub

		#End Region

		#Region "Export using FlexCelImageExport - MultiPageTiff"
		'How to create a multipage tiff using FlexCelImgExport.        
		'This will create a multipage tiff with the data.
		Private Sub DoExportMultiPageTiff(ByVal ColorDepth As ImageColorDepth, ByVal IsFax As Boolean)
			If Not HasFileOpen() Then
				Return
			End If
			If Not LoadPreferences() Then
				Return
			End If

			If exportTiffDialog.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
				Return
			End If

			Dim ExportType As ImageExportType = ImageExportType.Tiff
			If IsFax Then
				ExportType = ImageExportType.Fax
			End If

			Using ImgExport As New FlexCelImgExport(flexCelPrintDocument1.Workbook)
				ImgExport.AllVisibleSheets = cbAllSheets.Checked
				ImgExport.ResetPageNumberOnEachSheet = cbResetPageNumber.Checked

				ImgExport.Resolution = 96 'To get a better quality image but with larger file size too, increate this value. (for example to 300 or 600 dpi)
				Using TiffStream As New FileStream(exportTiffDialog.FileName, FileMode.Create)
					ImgExport.SaveAsImage(TiffStream, ExportType, ColorDepth)
				End Using
			End Using
			If MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
				Process.Start(exportTiffDialog.FileName)
			End If

		End Sub
		#End Region

		#Region "Event handlers"
		Private Sub ImgBlackAndWhite_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles blackAndWhiteToolStripMenuItem.Click
			DoExportUsingFlexCelImgExportComplex(ImageColorDepth.BlackAndWhite)
		End Sub

		Private Sub Img256Colors_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles colorsToolStripMenuItem.Click
			DoExportUsingFlexCelImgExportComplex(ImageColorDepth.Color256)
		End Sub

		Private Sub ImgTrueColor_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles trueColorToolStripMenuItem.Click
			DoExportUsingFlexCelImgExportComplex(ImageColorDepth.TrueColor)
		End Sub

		Private Sub ImgBlackAndWhite2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles blackAndWhiteToolStripMenuItem1.Click
			DoExportUsingFlexCelImgExportSimple(ImageColorDepth.BlackAndWhite)
		End Sub

		Private Sub Img256Colors2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles colorsToolStripMenuItem1.Click
			DoExportUsingFlexCelImgExportSimple(ImageColorDepth.Color256)
		End Sub

		Private Sub ImgTrueColor2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles trueColorToolStripMenuItem1.Click
			DoExportUsingFlexCelImgExportSimple(ImageColorDepth.TrueColor)
		End Sub

		Private Sub TiffFax_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles faxToolStripMenuItem.Click
			DoExportMultiPageTiff(ImageColorDepth.BlackAndWhite, True)
		End Sub

		Private Sub TiffBlackAndWhite_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles blackAndWhiteToolStripMenuItem2.Click
			DoExportMultiPageTiff(ImageColorDepth.BlackAndWhite, False)
		End Sub

		Private Sub Tiff256Colors_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles colorsToolStripMenuItem2.Click
			DoExportMultiPageTiff(ImageColorDepth.Color256, False)
		End Sub

		Private Sub TiffTrueColor_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles trueColorToolStripMenuItem2.Click
			DoExportMultiPageTiff(ImageColorDepth.TrueColor, False)
		End Sub

		#End Region

		Private Sub cbAllSheets_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbAllSheets.CheckedChanged
			cbSheet.Enabled = Not cbAllSheets.Checked
			cbResetPageNumber.Enabled = cbAllSheets.Checked
			Landscape.Enabled = Not cbAllSheets.Checked 'When exporting many sheets, we will honor the landscape/portrait setting on each one.
		End Sub



		#End Region

	End Class
End Namespace
