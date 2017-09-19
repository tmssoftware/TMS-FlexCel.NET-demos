Imports System.Collections
Imports System.ComponentModel
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports System.IO
Imports System.Reflection

Namespace HyperLinks
	''' <summary>
	''' How to deal with Hyperlinks in FlexCel.
	''' </summary>
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

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

		Private Sub button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.Click
			Close()
		End Sub

		Private Xls As ExcelFile = Nothing

		Private Sub ReadHyperLinks_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReadHyperlinks.Click
			If openFileDialog1.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
				Return
			End If
			Xls = New XlsFile()

			Xls.Open(openFileDialog1.FileName)

			dataGrid.CaptionText = "Hyperlinks on file: " & openFileDialog1.FileName
			HlDataTable.Rows.Clear()


			For i As Integer = 1 To Xls.HyperLinkCount
				Dim Range As TXlsCellRange = Xls.GetHyperLinkCellRange(i)
				Dim HLink As THyperLink = Xls.GetHyperLink(i)

				Dim HLinkType As String = System.Enum.GetName(GetType(THyperLinkType), HLink.LinkType)

				Dim values() As Object ={i, TCellAddress.EncodeColumn(Range.Left) &Range.Top.ToString(), TCellAddress.EncodeColumn(Range.Right) &Range.Bottom.ToString(), HLinkType, HLink.Text, HLink.Description, HLink.TextMark, HLink.TargetFrame, HLink.Hint }
				HlDataTable.Rows.Add(values)

			Next i

		End Sub

		Private Sub writeHyperLinks_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnWriteHyperlinks.Click
			If Xls Is Nothing Then
				MessageBox.Show("You need to open a file first.")
				Return
			End If

			Dim XlsOut As ExcelFile = New XlsFile(True)
			XlsOut.NewFile(1)

			For i As Integer = 1 To Xls.HyperLinkCount
				Dim Range As TXlsCellRange = Xls.GetHyperLinkCellRange(i)
				Dim HLink As THyperLink = Xls.GetHyperLink(i)

				Dim XF As Integer = -1
				Dim Value As Object = Xls.GetCellValue(Range.Top, Range.Left, XF)
				XlsOut.SetCellValue(i, 1, Value, XlsOut.AddFormat(Xls.GetFormat(XF)))
				XlsOut.AddHyperLink(New TXlsCellRange(i, 1, i, 1), HLink)
			Next i

			If saveFileDialog1.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
				Return
			End If
			XlsOut.Save(saveFileDialog1.FileName)
			If MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
				Process.Start(saveFileDialog1.FileName)
			End If
		End Sub

	End Class
End Namespace
