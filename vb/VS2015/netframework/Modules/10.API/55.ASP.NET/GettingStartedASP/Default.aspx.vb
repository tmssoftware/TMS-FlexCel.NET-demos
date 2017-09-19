Imports System
Imports System.Data
Imports System.Configuration
Imports System.Web
Imports System.Web.Security
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.WebControls.WebParts
Imports System.Web.UI.HtmlControls

Imports System.IO
Imports System.Reflection
Imports System.Drawing

Imports FlexCel.Core
Imports FlexCel.Render
Imports FlexCel.XlsAdapter

Partial Public Class _Default
	Inherits System.Web.UI.Page

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Private Sub CreateFile(ByVal Xls As ExcelFile)
		'Create a new file. We could also open an existing file with Xls.Open
		Xls.NewFile(1)
		'Set some cell values.
		Xls.SetCellValue(1, 1, "Hello to everybody")
		Xls.SetCellValue(2, 1, 3)
		Xls.SetCellValue(3, 1, 2.1)
		Xls.SetCellValue(4, 1, New TFormula("=Sum(A2,A3)"))

		'Load an image from disk.
		Dim AssemblyPath As String = HttpContext.Current.Request.PhysicalApplicationPath
		Using Img As System.Drawing.Image = System.Drawing.Image.FromFile(Path.Combine(Path.Combine(AssemblyPath, "images"), "Test.bmp"))

			'Add a new image on cell E5
			Xls.AddImage(2, 6, Img)
			'Add a new image with custom properties at cell F6
			Xls.AddImage(Img, New TImageProperties(New TClientAnchor(TFlxAnchorType.DontMoveAndDontResize, 2, 10, 6, 10, 100, 100, Xls), ""))
			'Swap the order of the images. it is not really necessary here, we could have loaded them on the inverse order.
			Xls.BringToFront(1)
		End Using

		'Add a comment on cell a2
		Xls.SetComment(2, 1, "This is 3")

		'Custom Format cells a2 and a3
		Dim f As TFlxFormat = Xls.GetDefaultFormat
		f.Font.Name = "Times New Roman"
		f.Font.Color = Color.Red
		f.FillPattern.Pattern = TFlxPatternStyle.LightDown
		f.FillPattern.FgColor = Color.Blue
		f.FillPattern.BgColor = Color.White

		Dim XF As Integer = Xls.AddFormat(f)

		Xls.SetCellFormat(2, 1, XF)
		Xls.SetCellFormat(3, 1, XF)

		f.Rotation = 45
		f.FillPattern.Pattern = TFlxPatternStyle.Solid
		Dim XF2 As Integer = Xls.AddFormat(f)
		'Apply a custom format to all the row.
		Xls.SetRowFormat(1, XF2)

		'Merge cells
		Xls.MergeCells(5, 1, 10, 6)
		'Note how this one merges with the previous range, creating a final range (5,1,15,6)
		Xls.MergeCells(10, 6, 15, 6)

		'Make sure rows are autofitted for pdf export.
		 Xls.AutofitRowsOnWorkbook(False, True, 1)

	End Sub


	Protected Sub BtnReadCellA1_Click(ByVal sender As Object, ByVal e As EventArgs)
		Dim Xls As ExcelFile = New XlsFile()
		If FileBox.PostedFile Is Nothing OrElse FileBox.PostedFile.InputStream Is Nothing OrElse FileBox.PostedFile.InputStream.Length = 0 Then

			LabelA1.Text = "No file selected"
			Return
		End If
		FileBox.PostedFile.InputStream.Position = 0
		Try
			Xls.Open(FileBox.PostedFile.InputStream)
			Dim v As Object = Xls.GetCellValue(1, 1)
			If v Is Nothing Then
				LabelA1.Text = "Cell A1 is empty"
			Else
				LabelA1.Text = "Cell A1 has the value: " & Convert.ToString(v)
			End If
		Catch ex As Exception
			LabelA1.Text = ex.Message
		End Try

	End Sub
	Protected Sub BtnXls_Click(ByVal sender As Object, ByVal e As EventArgs)
		Dim Xls As ExcelFile = New XlsFile()
		CreateFile(Xls)
		Using ms As New MemoryStream()
			Xls.Save(ms)
			ms.Position = 0
			Response.Clear()
			Response.AddHeader("Content-Disposition", "attachment; filename=Test.xls")
			Response.AddHeader("Content-Length", ms.Length.ToString())
			Response.ContentType = "application/excel" 'octet-stream";
			Response.BinaryWrite(ms.ToArray())
			Response.End()
		End Using
	End Sub
	Protected Sub BtnPdf_Click(ByVal sender As Object, ByVal e As EventArgs)
		Dim Xls As ExcelFile = New XlsFile()
		CreateFile(Xls)

		Dim Pdf As New FlexCelPdfExport(Xls)

		Using ms As New MemoryStream()

			Pdf.BeginExport(ms)
			Try
				Pdf.ExportAllVisibleSheets(True, "Getting Started")
			Finally
				Pdf.EndExport()
			End Try
			ms.Position = 0
			Response.Clear()
			Response.AddHeader("Content-Disposition", "attachment; filename=Test.pdf")
			Response.AddHeader("Content-Length", ms.Length.ToString())
			Response.ContentType = "application/pdf" 'octet-stream";
			Response.BinaryWrite(ms.ToArray())
			Response.End()
		End Using

	End Sub
End Class
