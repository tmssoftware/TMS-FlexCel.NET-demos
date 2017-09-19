Imports System
Imports System.Data
Imports System.Configuration
Imports System.Web
Imports System.Web.Security
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.WebControls.WebParts
Imports System.Web.UI.HtmlControls

Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Render
Imports FlexCel.AspNet

Partial Public Class _Default
	Inherits System.Web.UI.Page

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
		Dim xls As New XlsFile()

		Dim DefaultFile As String = Server.MapPath("~/default.xls")
		If IsPostBack Then
			If Uploader.HasFile Then
				xls.Open(Uploader.FileContent)
			Else
				xls.Open(DefaultFile)
			End If
		Else
			xls.Open(DefaultFile)
		End If

		Viewer.HtmlExport.ImageNaming = TImageNaming.Guid
		Viewer.HtmlExport.Workbook = xls
		Viewer.RelativeImagePath = "images"
		Viewer.HtmlExport.FixIE6TransparentPngSupport = True 'This is only needed if you are using IE and there are transparent png files.
		Viewer.ImageExportMode = TImageExportMode.TemporaryFiles

	End Sub
End Class
