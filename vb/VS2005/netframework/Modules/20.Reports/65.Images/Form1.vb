Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Report
Imports FlexCel.Demo.SharedData


Namespace Images
	''' <summary>
	''' A report with lots of images.
	''' </summary>
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles button1.Click
			AutoRun()
		End Sub

		Public Sub AutoRun()
			Using ordersReport As FlexCelReport = SharedData.CreateReport()
				AddHandler ordersReport.GetImageData, AddressOf ordersReport_GetImageData
				ordersReport.SetValue("Date", Date.Now)

				Dim DataPath As String = Path.Combine(Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), ".."), "..") & Path.DirectorySeparatorChar

				If saveFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
					ordersReport.Run(DataPath & "Images.template.xls", saveFileDialog1.FileName)

					If MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
						Process.Start(saveFileDialog1.FileName)
					End If
				End If
			End Using
		End Sub

		Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
			Close()
		End Sub

		Private Sub ordersReport_GetImageData(ByVal sender As Object, ByVal e As FlexCel.Report.GetImageDataEventArgs)
			If String.Compare(e.ImageName, "<#PhotoCode>", True) = 0 Then
				Dim RealImageData() As Byte = ImageUtils.StripOLEHeader(e.ImageData) 'On access databases, images are stored with an OLE
				'header that we have to strip to get the real image.
				'This is done automatically by flexcel in most cases,
				'but here we have the original image format.
				Using MemStream As New MemoryStream(RealImageData) 'Keep stream open until bitmap has been used
					Using bmp As New Bitmap(MemStream)
						bmp.RotateFlip(RotateFlipType.Rotate90FlipNone)
						Using OutStream As New MemoryStream()
							bmp.Save(OutStream, System.Drawing.Imaging.ImageFormat.Png)
							e.Width = bmp.Width
							e.Height = bmp.Height
							e.ImageData = OutStream.ToArray()
						End Using
					End Using
				End Using
			End If

		End Sub
	End Class

End Namespace
