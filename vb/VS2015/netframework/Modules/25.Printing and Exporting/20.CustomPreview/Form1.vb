Imports System.ComponentModel
Imports System.IO
Imports System.Drawing.Drawing2D

Imports System.Threading

Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Winforms
Imports FlexCel.Render
Imports FlexCel.Pdf

Namespace CustomPreview
	''' <summary>
	''' Previewer of files.
	''' </summary>
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Public Sub New()
			Me.New(New String(){})
		End Sub

		Public Sub New(ByVal Args() As String)
			InitializeComponent()
			ResizeToolbar(mainToolbar)
			If Args.Length > 0 Then
				LoadFile(Args(0))
			End If

			If ExcelFile.SupportsXlsx Then
				Me.openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|Excel 97/2003|*.xls|Excel 2007|*.xlsx;*.xlsm|All files|*.*"
			End If

			MainPreview.CenteredPreview = True
			thumbs.CenteredPreview = True
		End Sub

		Private Sub ResizeToolbar(ByVal toolbar As ToolStrip)

			Using gr As Graphics = CreateGraphics()
				Dim xFactor As Double = gr.DpiX / 96.0
				Dim yFactor As Double = gr.DpiY / 96.0
				toolbar.ImageScalingSize = New Size(CInt(Fix(24 * xFactor)), CInt(Fix(24 * yFactor)))
				toolbar.Width = 0 'force a recalc of the buttons.
			End Using
		End Sub


		Private Sub UpdatePages()
			edPage.Text = String.Format("{0} of {1}", MainPreview.StartPage, MainPreview.TotalPages)
		End Sub

		Private Sub flexCelPreview1_StartPageChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MainPreview.StartPageChanged
			UpdatePages()
		End Sub

		Private Sub ChangePages()
			Dim s As String = edPage.Text.Trim()
			Dim pos As Integer = 0
			Do While pos < s.Length AndAlso s.Chars(pos) >= "0"c AndAlso s.Chars(pos) <= "9"c
				pos += 1
			Loop
			If pos > 0 Then
				Dim page As Integer = MainPreview.StartPage
				Try
					page = Convert.ToInt32(s.Substring(0, pos))
				Catch e1 As Exception
				End Try

				MainPreview.StartPage = page
			End If
			UpdatePages()
		End Sub

		Private Sub edPage_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles edPage.Leave
			ChangePages()
		End Sub

		Private Sub edPage_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles edPage.KeyPress
			If e.KeyChar = ChrW(13) Then
				ChangePages()
			End If
			If e.KeyChar = ChrW(27) Then
				UpdatePages()
			End If
		End Sub

		Private Sub flexCelPreview1_ZoomChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MainPreview.ZoomChanged
			UpdateZoom()
		End Sub

		Private Sub UpdateZoom()
			edZoom.Text = String.Format("{0}%", CInt(Fix(Math.Round(MainPreview.Zoom * 100))))
			If MainPreview.AutofitPreview = TAutofitPreview.None Then
				UpdateAutofitText()
			End If
		End Sub

		Private Sub ChangeZoom()
			Dim s As String = edZoom.Text.Trim()
			Dim pos As Integer = 0
			Do While pos < s.Length AndAlso s.Chars(pos) >= "0"c AndAlso s.Chars(pos) <= "9"c
				pos += 1
			Loop
			If pos > 0 Then
				Dim zoom As Integer = CInt(Fix(Math.Round(MainPreview.Zoom * 100)))
				Try
					zoom = Convert.ToInt32(s.Substring(0, pos))
				Catch e1 As Exception
				End Try

				MainPreview.Zoom = zoom / 100.0
			End If
			UpdateZoom()
		End Sub

		Private Sub edZoom_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles edZoom.KeyPress
			If e.KeyChar = ChrW(13) Then
				ChangeZoom()
			End If
			If e.KeyChar = ChrW(27) Then
				UpdateZoom()
			End If
		End Sub

		Private Sub edZoom_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles edZoom.Enter
			ChangeZoom()
		End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
			Close()
		End Sub

		Private Sub openFile_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles openFile.Click
			If openFileDialog.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
				Return
			End If
			LoadFile(openFileDialog.FileName)
		End Sub

		'The event that will actually provide the password to open the empty form.
		Private Sub GetPassword(ByVal e As OnPasswordEventArgs)
			Dim Pwd As New PasswordForm()
			e.Password = String.Empty
			If Pwd.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
				Return
			End If
			e.Password = Pwd.Password
		End Sub


		Friend Sub LoadFile(ByVal FileName As String)
			openFileDialog.FileName = FileName
			lbSheets.Items.Clear()

			Dim xls As New XlsFile()
			xls.HeadingColWidth = -1
			xls.HeadingRowHeight = -1
			AddHandler xls.Protection.OnPassword, AddressOf GetPassword
			xls.Open(FileName)

			For i As Integer = 1 To xls.SheetCount
				lbSheets.Items.Add(xls.GetSheetName(i))
			Next i

			lbSheets.SelectedIndex = xls.ActiveSheet - 1

			flexCelImgExport1.Workbook = xls
			MainPreview.InvalidatePreview()
			Text = "Custom Preview: " & openFileDialog.FileName
			'btnHeadings.Checked = flexCelImgExport1.Workbook.PrintHeadings;
			'btnGridLines.Checked = flexCelImgExport1.Workbook.PrintGridLines;
			btnFirst.Enabled = True
			btnPrev.Enabled = True
			btnNext.Enabled = True
			btnLast.Enabled = True
			edPage.Enabled = True
			btnZoomIn.Enabled = True
			edZoom.Enabled = True
			btnZoomOut.Enabled = True
			btnGridLines.Enabled = True
			btnHeadings.Enabled = True
			btnRecalc.Enabled = True
			btnPdf.Enabled = True

		End Sub

		Private Sub btnFirst_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFirst.Click
			MainPreview.StartPage = 1
		End Sub

		Private Sub btnPrev_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrev.Click
			MainPreview.StartPage -= 1
		End Sub

		Private Sub btnNext_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNext.Click
			MainPreview.StartPage += 1
		End Sub

		Private Sub btnLast_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLast.Click
			MainPreview.StartPage = MainPreview.TotalPages
		End Sub

		Private Sub btnZoomOut_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnZoomOut.Click
			MainPreview.Zoom -= 0.1
		End Sub

		Private Sub btnZoomIn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnZoomIn.Click
			MainPreview.Zoom += 0.1
		End Sub

		Private Sub lbSheets_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbSheets.SelectedIndexChanged
			If flexCelImgExport1.Workbook Is Nothing Then
				Return
			End If
			If lbSheets.Items.Count > flexCelImgExport1.Workbook.SheetCount Then
				Return
			End If
			flexCelImgExport1.Workbook.ActiveSheet = lbSheets.SelectedIndex + 1
			MainPreview.InvalidatePreview()
		End Sub

		Private Sub btnPdf_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPdf.Click
			If flexCelImgExport1.Workbook Is Nothing Then
				MessageBox.Show("There is no open file")
				Return
			End If
			If PdfSaveFileDialog.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
				Return
			End If

			Using PdfExport As New FlexCelPdfExport(flexCelImgExport1.Workbook, True)
				If Not DoExportToPdf(PdfExport) Then
					Return
				End If
			End Using

			If MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> System.Windows.Forms.DialogResult.Yes Then
				Return
			End If
			Process.Start(PdfSaveFileDialog.FileName)
		End Sub

		Private Function DoExportToPdf(ByVal PdfExport As FlexCelPdfExport) As Boolean
			Dim MyPdfThread As New PdfThread(PdfExport, PdfSaveFileDialog.FileName, cbAllSheets.Checked)
			Dim PdfExportThread As New Thread(New ThreadStart(AddressOf MyPdfThread.ExportToPdf))
			PdfExportThread.Start()
			Using Pg As New PdfProgressDialog()
				Pg.ShowProgress(PdfExportThread, PdfExport)
				If Pg.DialogResult <> System.Windows.Forms.DialogResult.OK Then
					PdfExport.Cancel()
					PdfExportThread.Join() 'We could just leave the thread running until it dies, but there are 2 reasons for waiting until it finishes:
											'1) We could dispose it before it ends. This is workaroundable.
											'2) We might change its workbook object before it ends (by loading other file). This will surely bring issues.
					Return False
				End If

				If MyPdfThread IsNot Nothing AndAlso MyPdfThread.MainException IsNot Nothing Then
					Throw MyPdfThread.MainException
				End If
			End Using
			Return True
		End Function

		Private Sub cbAllSheets_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbAllSheets.CheckedChanged
			lbSheets.Visible = Not cbAllSheets.Checked
			sheetSplitter.Visible = lbSheets.Visible
			flexCelImgExport1.AllVisibleSheets = cbAllSheets.Checked
			If flexCelImgExport1.Workbook Is Nothing Then
				Return
			End If
			MainPreview.InvalidatePreview()

		End Sub

		Private Sub btnRecalc_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRecalc.Click
			If flexCelImgExport1.Workbook Is Nothing Then
				MessageBox.Show("Please open a file before recalculating.")
				Return
			End If
			flexCelImgExport1.Workbook.Recalc(True)
			MainPreview.InvalidatePreview()

		End Sub


		Private Sub mainForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
		End Sub

		Private Sub btnHeadings_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnHeadings.Click
			Dim xls As ExcelFile = flexCelImgExport1.Workbook
			If xls Is Nothing Then
				Return
			End If

			If cbAllSheets.Checked Then
				Dim SaveActiveSheet As Integer = xls.ActiveSheet
				For sheet As Integer = 1 To xls.SheetCount
					xls.ActiveSheet = sheet
					xls.PrintHeadings = btnHeadings.Checked
				Next sheet
				xls.ActiveSheet = SaveActiveSheet
			Else
				xls.PrintHeadings = btnHeadings.Checked
			End If
			MainPreview.InvalidatePreview()

		End Sub

		Private Sub btnGridLines_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnGridLines.Click
			Dim xls As ExcelFile = flexCelImgExport1.Workbook
			If xls Is Nothing Then
				Return
			End If

			If cbAllSheets.Checked Then
				Dim SaveActiveSheet As Integer = xls.ActiveSheet
				For sheet As Integer = 1 To xls.SheetCount
					xls.ActiveSheet = sheet
					xls.PrintGridLines = btnGridLines.Checked
				Next sheet
				xls.ActiveSheet = SaveActiveSheet
			Else
				xls.PrintGridLines = btnGridLines.Checked
			End If
			MainPreview.InvalidatePreview()

		End Sub

		Private Sub noneToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles noneToolStripMenuItem.Click
			MainPreview.AutofitPreview = TAutofitPreview.None
			UpdateAutofitText()
		End Sub

		Private Sub fitToWidthToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles fitToWidthToolStripMenuItem.Click
			MainPreview.AutofitPreview = TAutofitPreview.Width
			UpdateAutofitText()
		End Sub

		Private Sub fitToHeightToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles fitToHeightToolStripMenuItem.Click
			MainPreview.AutofitPreview = TAutofitPreview.Height
			UpdateAutofitText()
		End Sub

		Private Sub fitToPageToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles fitToPageToolStripMenuItem.Click
			MainPreview.AutofitPreview = TAutofitPreview.Full
			UpdateAutofitText()
		End Sub

		Private Sub UpdateAutofitText()
			Select Case MainPreview.AutofitPreview
				Case TAutofitPreview.None
					btnAutofit.Text = "No Autofit"
				Case TAutofitPreview.Width
					btnAutofit.Text = "Fit to Width"
				Case TAutofitPreview.Height
					btnAutofit.Text = "Fit to Height"
				Case TAutofitPreview.Full
					btnAutofit.Text = "Fit to Page"
				Case Else
			End Select

		End Sub

	End Class

	#Region "PdfThread"
	Friend Class PdfThread
		Private PdfExport As FlexCelPdfExport
		Private FileName As String
		Private AllVisibleSheets As Boolean
		Private FMainException As Exception

		Friend Sub New(ByVal aPdfExport As FlexCelPdfExport, ByVal aFileName As String, ByVal aAllVisibleSheets As Boolean)
			PdfExport = aPdfExport
			FileName = aFileName
			AllVisibleSheets = aAllVisibleSheets
		End Sub

		Friend Sub ExportToPdf()
			Try
				If AllVisibleSheets Then
					Try
						Using f As New FileStream(FileName, FileMode.Create, FileAccess.Write)
							PdfExport.BeginExport(f)
							PdfExport.PageLayout = TPageLayout.Outlines
							PdfExport.ExportAllVisibleSheets(False, System.IO.Path.GetFileNameWithoutExtension(FileName))
							PdfExport.EndExport()
						End Using
					Catch
						Try
							File.Delete(FileName)
						Catch
							'Not here.
						End Try
						Throw
					End Try
				Else
					PdfExport.PageLayout = TPageLayout.None
					PdfExport.Export(FileName)
				End If
			Catch ex As Exception
				FMainException = ex
			End Try
		End Sub

		Friend ReadOnly Property MainException() As Exception
			Get
				Return FMainException
			End Get
		End Property
	End Class
	#End Region


End Namespace
