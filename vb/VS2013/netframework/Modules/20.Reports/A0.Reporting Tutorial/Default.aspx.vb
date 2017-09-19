Imports System
Imports System.Collections
Imports System.Configuration
Imports System.Data
Imports System.Web
Imports System.Web.Security
Imports System.Web.UI
Imports System.Web.UI.HtmlControls
Imports System.Web.UI.WebControls
Imports System.Web.UI.WebControls.WebParts
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Render
Imports FlexCel.Report
Imports System.IO

Partial Public Class _Default
	Inherits System.Web.UI.Page

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Private Function CreateReport() As ExcelFile
		Dim Result As New XlsFile(True)
		Result.Open(MapPath("~/App_Data/template.xls"))
		Using fr As New FlexCelReport()
			LoadData(fr)

			fr.SetValue("ReportCaption", "Hello from FlexCel!")
			fr.Run(Result)
			Return Result
		End Using
	End Function

	Private Sub LoadData(ByVal fr As FlexCelReport)
		Dim Data As New DataSet1()
		Dim ProductAdapter As New DataSet1TableAdapters.ProductTableAdapter()
		ProductAdapter.Fill(Data.Product)

		Dim ProductPhotoAdapter As New DataSet1TableAdapters.ProductPhotoTableAdapter()
		ProductPhotoAdapter.Fill(Data.ProductPhoto)

		Dim ProductProductPhotoAdapter As New DataSet1TableAdapters.ProductProductPhotoTableAdapter()
		ProductProductPhotoAdapter.Fill(Data.ProductProductPhoto)

		fr.AddTable(Data)
	End Sub

	Protected Sub Button1_Click(ByVal sender As Object, ByVal e As EventArgs)
		Dim xls As ExcelFile = CreateReport()
		FlexCelAspViewer1.HtmlExport.Workbook = xls
	End Sub
	Protected Sub Button2_Click(ByVal sender As Object, ByVal e As EventArgs)
		Dim xls As ExcelFile = CreateReport()

		Using ms As New MemoryStream()
			xls.Save(ms)
			ms.Position = 0
			Response.Clear()
			Response.AddHeader("Content-Disposition", "attachment; filename=Test.xls")
			Response.AddHeader("Content-Length", ms.Length.ToString())
			Response.ContentType = "application/excel" 'octet-stream";
			Response.BinaryWrite(ms.ToArray())
			Response.End()
		End Using


	End Sub
	Protected Sub Button3_Click(ByVal sender As Object, ByVal e As EventArgs)
		Dim xls As ExcelFile = CreateReport()

		Using ms As New MemoryStream()
			Using pdf As New FlexCelPdfExport()
				pdf.Workbook = xls
				pdf.BeginExport(ms)
				pdf.ExportAllVisibleSheets(False, "FlexCel")
				pdf.EndExport()
				ms.Position = 0
				Response.Clear()
				Response.AddHeader("Content-Disposition", "attachment; filename=Test.pdf")
				Response.AddHeader("Content-Length", ms.Length.ToString())
				Response.ContentType = "application/pdf" 'octet-stream";
				Response.BinaryWrite(ms.ToArray())
				Response.End()
			End Using
		End Using

	End Sub
End Class

