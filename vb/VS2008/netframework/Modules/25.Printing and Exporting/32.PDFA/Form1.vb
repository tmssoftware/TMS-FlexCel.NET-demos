Imports System.Collections
Imports System.ComponentModel
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports System.IO
Imports System.Reflection
Imports System.Drawing.Drawing2D
Imports FlexCel.Pdf

Imports FlexCel.Render


Namespace PDFA
	''' <summary>
	''' Exporting xls files to PDF/A.
	''' </summary>
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Public Sub New()
			InitializeComponent()
			cbPdfType.SelectedIndex = 1
		End Sub

		Private Sub button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
			Close()
		End Sub

		Private Sub export_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles export.Click
			Dim EmbedSource As Boolean = cbEmbed.Checked
			Dim PdfType As TPdfType = GetPdfType()
			Dim TagMode As TTagMode = GetTagMode()

			If EmbedSource Then
				If PdfType <> TPdfType.PDFA3 AndAlso PdfType <> TPdfType.Standard Then
					MessageBox.Show("To embed a file, you need to use standard PDF or PDF/A3")
					Return
				End If
			End If

			If exportDialog.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
				Return
			End If

			CreateFile(exportDialog.FileName, EmbedSource, PdfType, TagMode)

			If MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
				Process.Start(exportDialog.FileName)
			End If

		End Sub

		Private Function GetPdfType() As TPdfType
			Select Case cbPdfType.SelectedIndex
				Case 0
					Return TPdfType.Standard
				Case 1, 2
					Return TPdfType.PDFA1
				Case 3, 4
					Return TPdfType.PDFA2
				Case 5, 6
					Return TPdfType.PDFA3
			End Select

			Throw New Exception("Unexpected PDF type")
		End Function

		Private Function GetTagMode() As TTagMode
			Select Case cbPdfType.SelectedIndex
				Case 0, 1, 3, 5
					Return TTagMode.Full
			End Select
			Return TTagMode.None
		End Function

		Private Sub CreateFile(ByVal FileName As String, ByVal EmbedSource As Boolean, ByVal PdfType As TPdfType, ByVal TagMode As TTagMode)
			Dim xls As ExcelFile = CreateSourceFile()
			Using pdf As New FlexCelPdfExport(xls, True)
				pdf.PdfType = PdfType
				pdf.TagMode = TagMode
				If EmbedSource Then
                    pdf.AttachFile("Report.xlsx", StandardMimeType.Xlsx, "This is the source file used to create the PDF", Date.Now,  TPdfAttachmentKind.Source, _
                       AddressOf New AttachWriter(xls).SaveAttachment)
				End If
				pdf.Export(FileName)
			End Using
		End Sub

		Private Function CreateSourceFile() As ExcelFile
			Dim xls As ExcelFile = New XlsFile()
			xls.NewFile(1, TExcelFileFormat.v2019)
			xls.SetCellValue(1, 1, "This is a test from FlexCel!")
			xls.SetCellValue(2, 1, "Here is some emoji to show unicode surrogate support: 🐜🐏")
			xls.SetCellValue(3, 1, "You might need a font able to show emoji for those characters to show")
			xls.SetCellValue(4, 1, "Windows 7 and 8 have SegoeUISymbol, which can show them and is used automatically by FlexCel.")
			Return xls
		End Function

    End Class

                 Class AttachWriter
                     Private xls As ExcelFile

                     Public Sub New(axls As ExcelFile)
                         xls = axls
                    End Sub

                     Public Sub SaveAttachment(attachWriter As TPdfAttachmentWriter)
                         Using ms As New MemoryStream()
                             xls.Save(ms, TFileFormats.Xlsx)
                             ms.Position = 0
                             attachWriter.Write(ms)
                         End Using
                     End Sub
                 End Class
End Namespace
