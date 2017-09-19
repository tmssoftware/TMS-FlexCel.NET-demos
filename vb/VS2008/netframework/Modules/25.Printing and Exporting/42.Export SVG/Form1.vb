Imports System.Collections
Imports System.ComponentModel
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Render
Imports System.IO

Imports System.Text



Namespace ExportSVG
	''' <summary>
	''' An Example on how to export to SVG.
	''' </summary>
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Private SVG As New FlexCelSVGExport()
		Public Sub New()
			InitializeComponent()
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

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
			Close()
		End Sub

		Private Sub LoadSheetConfig()
			Dim Xls As ExcelFile = SVG.Workbook

			chGridLines.Checked = Xls.PrintGridLines
			chPrintHeadings.Checked = Xls.PrintHeadings
			chFormulaText.Checked = Xls.ShowFormulaText
		End Sub

		Private Sub openFile_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles openFile.Click
			If openFileDialog1.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
				Return
			End If
			SVG.Workbook = New XlsFile()

			SVG.Workbook.Open(openFileDialog1.FileName)

			Text = "Export: " & openFileDialog1.FileName

			Dim Xls As ExcelFile = SVG.Workbook

			cbSheet.Items.Clear()
			For i As Integer = 1 To Xls.SheetCount
				cbSheet.Items.Add(Xls.GetSheetName(i))
			Next i
			cbSheet.SelectedIndex = Xls.ActiveSheet - 1

			LoadSheetConfig()
		End Sub

		Private Function CheckFileOpen() As Boolean
			If SVG.Workbook Is Nothing Then
				MessageBox.Show("You need to open a file first.")
				Return False
			End If
			Return True
		End Function

		Private Function LoadPreferences() As Boolean
			'NOTE: THERE SHOULD BE *A LOT* MORE VALIDATION OF VALUES ON THIS METHOD. (For example, validate that margins are between bounds)
			' As this is a simple demo, they are not included. 
			Try
				Dim Xls As ExcelFile = SVG.Workbook

				'Note: In this demo we will only apply this things to the active sheet.
				'If you want to apply the settings to all the sheets, you should loop in the sheets and change them here.
				Xls.PrintGridLines = chGridLines.Checked
				Xls.PrintHeadings = chPrintHeadings.Checked
				Xls.ShowFormulaText = chFormulaText.Checked

				SVG.PrintRangeLeft = Convert.ToInt32(edLeft.Text)
				SVG.PrintRangeTop = Convert.ToInt32(edTop.Text)
				SVG.PrintRangeRight = Convert.ToInt32(edRight.Text)
				SVG.PrintRangeBottom = Convert.ToInt32(edBottom.Text)

				SVG.HidePrintObjects = THidePrintObjects.None
				If Not cbImages.Checked Then
					SVG.HidePrintObjects = SVG.HidePrintObjects Or THidePrintObjects.Images
				End If
				If Not cbHyperlinks.Checked Then
					SVG.HidePrintObjects = SVG.HidePrintObjects Or THidePrintObjects.Hyperlynks
				End If
				If Not cbComments.Checked Then
					SVG.HidePrintObjects = SVG.HidePrintObjects Or THidePrintObjects.Comments
				End If
				If Not cbHeadersFooters.Checked Then
					SVG.HidePrintObjects = SVG.HidePrintObjects Or THidePrintObjects.HeadersAndFooters
				End If

			Catch e As Exception
				MessageBox.Show("Error: " & e.Message)
				Return False
			End Try
			Return True
		End Function


		Private Sub cbSheet_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
			SVG.Workbook.ActiveSheet = cbSheet.SelectedIndex + 1
			LoadSheetConfig()
		End Sub

		Private Sub export_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles export.Click
			If Not CheckFileOpen() Then
				Return
			End If
			If Not LoadPreferences() Then
				Return
			End If

			If exportDialog.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
				Return
			End If

			SVG.AllowOverwritingFiles = True

			SVG.AllVisibleSheets = cbExportObject.SelectedIndex = 0

                       SVG.SaveAsImage(AddressOf SaveSVG)

 

			If MessageBox.Show("Do you want to open the folder with the generated files?", "Confirm", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
				Process.Start(Path.GetDirectoryName(exportDialog.FileName))
			End If

		End Sub


		Private Sub mainForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
			cbExportObject.SelectedIndex = 0
		End Sub

		Private Sub cbExportObject_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbExportObject.SelectedIndexChanged
			cbSheet.Enabled = cbExportObject.SelectedIndex = 1
		End Sub
               Private Sub SaveSVG(x As SVGExportParameters)
                 x.FileName = Path.ChangeExtension(exportDialog.FileName, "") & "_" & x.Workbook.SheetName & "_" & x.SheetPageNumber.ToString() & ".svg"
               End Sub
	End Class
End Namespace
