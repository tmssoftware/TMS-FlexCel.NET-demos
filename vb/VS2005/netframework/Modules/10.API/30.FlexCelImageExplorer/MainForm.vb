Imports System.Collections
Imports System.ComponentModel
Imports System.IO

Imports FlexCel.Core
Imports FlexCel.XlsAdapter

Namespace FlexCelImageExplorer
	''' <summary>
	''' Image Explorer.
	''' </summary>
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Public Sub New()
			InitializeComponent()
			AddHandler GetCurrencyManager.CurrentChanged, AddressOf CurrentRowChanged
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

		Private CurrentFilename As String = Nothing
		Private CompressForm As TCompressForm

		Private ReadOnly Property GetCurrencyManager() As CurrencyManager
			Get
				Return CType(Me.BindingContext(dataGrid.DataSource, dataGrid.DataMember), CurrencyManager)
			End Get
		End Property

		Private ReadOnly Property GetImagePos() As Integer
			Get
				Return GetCurrencyManager.Position
			End Get
		End Property

		Private Sub btnExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.Click
			Close()
		End Sub

		Private Sub btnOpenFile_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOpenFile.Click
			If openFileDialog1.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
				Return
			End If
			OpenFile(openFileDialog1.FileName)
			If cbScanFolder.Checked Then
				FillListBox()
			End If
		End Sub

		Private Function GetHasCrop(ByVal ImgProps As TImageProperties) As Boolean
			Return ImgProps.CropArea.CropFromLeft <> 0 OrElse ImgProps.CropArea.CropFromRight <> 0 OrElse ImgProps.CropArea.CropFromTop <> 0 OrElse ImgProps.CropArea.CropFromBottom <> 0
		End Function

		Private Sub FillListBox()
			lblFolder.Text = "Files on folder: " & Path.GetDirectoryName(openFileDialog1.FileName)
			Dim di As New DirectoryInfo(Path.GetDirectoryName(openFileDialog1.FileName))
			Dim Fi() As FileInfo = di.GetFiles("*.xls")
			FilesListBox.Items.Clear()

			Dim Files(Fi.Length - 1) As TImageInfo

			For k As Integer = 0 To Fi.Length - 1
				Dim f As FileInfo = Fi(k)
				Dim HasCrop As Boolean = False
				Dim HasARGB As Boolean = False
				Dim x1 As New XlsFile()

				Dim HasImages As Boolean = False

				Try
					x1.Open(f.FullName)
					For sheet As Integer = 1 To x1.SheetCount
						x1.ActiveSheet = sheet
						For i As Integer = x1.ImageCount To 1 Step -1
							HasImages = True
							Dim ip As TImageProperties = x1.GetImageProperties(i)
							If Not HasCrop Then
								HasCrop = GetHasCrop(ip)
							End If

							Dim imgType As TXlsImgType = TXlsImgType.Unknown
							Using ms As New MemoryStream()
								x1.GetImage(i, imgType, ms)
								Dim PngInfo As FlexCel.Pdf.TPngInformation = FlexCel.Pdf.TPdfPng.GetPngInfo(ms)
								If PngInfo IsNot Nothing Then
									HasARGB = PngInfo.ColorType = 6
								End If
							End Using

						Next i
					Next sheet
				Catch e1 As Exception
					Files(k) = New TImageInfo(f, False, False, False, False)
					Continue For
				End Try

				Files(k) = New TImageInfo(f, True, HasCrop, HasImages, HasARGB)
			Next k

			FilesListBox.Items.AddRange(Files)
		End Sub

		Private Sub OpenFile(ByVal FileName As String)
			ImageDataTable.Rows.Clear()

			Try
				Dim Xls As New XlsFile(True)
				CurrentFilename = FileName
				Xls.Open(FileName)

				For sheet As Integer = 1 To Xls.SheetCount
					Xls.ActiveSheet = sheet
					For i As Integer = Xls.ImageCount To 1 Step -1
						Dim ImageType As TXlsImgType = TXlsImgType.Unknown
						Dim ImgBytes() As Byte = Xls.GetImage(i, ImageType)
						Dim ImgProps As TImageProperties = Xls.GetImageProperties(i)
						Dim ImgData(ImageDataTable.Columns.Count - 1) As Object
						ImgData(0) = Xls.SheetName
						ImgData(1) = i
						ImgData(4) = ImageType.ToString()
						ImgData(7) = Xls.GetImageName(i)
						ImgData(8) = ImgBytes
						ImgData(9) = GetHasCrop(ImgProps)


						Using ms As New MemoryStream(ImgBytes)
							Dim PngInfo As FlexCel.Pdf.TPngInformation = FlexCel.Pdf.TPdfPng.GetPngInfo(ms)
							If PngInfo IsNot Nothing Then
								ImgData(2) = PngInfo.Width
								ImgData(3) = PngInfo.Height
								Dim s As String = String.Empty
								Dim bpp As Integer = 0

								If (PngInfo.ColorType And 4) <> 0 Then
									s &= "ALPHA-"
									bpp = 1
								End If
								If (PngInfo.ColorType And 2) = 0 Then
									s &= "Grayscale -" & (1 << PngInfo.BitDepth).ToString() & " shades. "
									bpp = 1
								Else
									If (PngInfo.ColorType And 1) = 0 Then
										bpp += 3
										s &= "RGB - " & (PngInfo.BitDepth * (bpp)).ToString() & "bpp.  "
									Else
										s &= "Indexed - " & (1 << PngInfo.BitDepth).ToString() & " colors. "
										bpp = 1
									End If
								End If

								ImgData(5) = s

								ImgData(6) = (Math.Round(PngInfo.Width * PngInfo.Height * PngInfo.BitDepth * bpp / 8F / 1024F)).ToString() & " kb."
							Else
								ms.Position = 0
								Try
									Using Img As Image = Image.FromStream(ms)
										Dim Bmp As Bitmap = TryCast(Img, Bitmap)
										If Bmp IsNot Nothing Then
											ImgData(5) = Bmp.PixelFormat.ToString() & "bpp"
										End If
										ImgData(2) = Img.Width
										ImgData(3) = Img.Height
									End Using
								Catch e1 As Exception
									ImgData(2) = -1
									ImgData(3) = -1
									ImgData(5) = Nothing
									ImgData(8) = Nothing

								End Try
							End If
						End Using


						ImageDataTable.Rows.Add(ImgData)
					Next i
				Next sheet

			Catch ex As Exception
				MessageBox.Show(ex.Message, "Error")
				dataGrid.CaptionText = "No file selected"
				CurrentFilename = Nothing
				Return
			End Try
			dataGrid.CaptionText = "Selected file: " & FileName
			CurrentRowChanged(GetCurrencyManager, Nothing)
		End Sub

		Private Sub FilesListBox_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles FilesListBox.SelectedIndexChanged
			Dim ImageInfo As TImageInfo = CType(FilesListBox.SelectedItem, TImageInfo)
			If ImageInfo Is Nothing Then
				Return
			End If
			OpenFile(ImageInfo.File.FullName)
		End Sub

		Public Sub CurrentRowChanged(ByVal sender As Object, ByVal e As System.EventArgs)
			Dim Pos As Integer = CType(sender, BindingManagerBase).Position
			If Pos < 0 Then
				PreviewBox.Image = Nothing
				Return
			End If

			Dim r As DataRowCollection = dataSet1.Tables("ImageDataTable").Rows
			If Pos < 0 OrElse Pos >= r.Count Then
				PreviewBox.Image = Nothing
			Else
				Dim ImgData() As Byte = TryCast(r(Pos).ItemArray(8), Byte())
				If ImgData Is Nothing Then
					PreviewBox.Image = Nothing
				Else
					Using ms As New MemoryStream(ImgData)
						PreviewBox.Image = Image.FromStream(ms)
					End Using
				End If
			End If
		End Sub

		Private Sub btnOpen_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnShowInExcel.Click
			If CurrentFilename Is Nothing Then
				MessageBox.Show("There is no open file")
				Return
			End If
			System.Diagnostics.Process.Start(CurrentFilename)
		End Sub

		Private Sub btnConvert_Click(ByVal sender As Object, ByVal e As System.EventArgs)
			'This is not yet implemented...
			Dim Pos As Integer = GetImagePos
			If Pos < 0 Then
				MessageBox.Show("There is no selected image", "Error")
				Return
			End If
			If CompressForm Is Nothing Then
				CompressForm = New TCompressForm()
			End If
			CompressForm.ImageToUse = CType(dataSet1.Tables("ImageDataTable").Rows(Pos).ItemArray(8), Byte())
			CompressForm.XlsFilename = CurrentFilename
			CompressForm.ShowDialog()
		End Sub

		Private Sub FilesListBox_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles FilesListBox.DrawItem
			If e.Index < 0 Then
				Return
			End If
			e.DrawBackground()
			Dim myBrush As Brush = Brushes.Black

			Dim ImageInfo As TImageInfo = CType(CType(sender, ListBox).Items(e.Index), TImageInfo)
			If Not ImageInfo.HasImages Then
				myBrush = Brushes.Silver
			End If
			If ImageInfo.HasCrop Then
				myBrush = Brushes.Red
			End If

			Dim NewStyle As FontStyle
			If ImageInfo.HasARGB Then
				NewStyle = FontStyle.Bold
			Else
				NewStyle = FontStyle.Regular
			End If
			Using MyFont As New Font(e.Font, NewStyle)
				e.Graphics.DrawString(ImageInfo.ToString(), MyFont, myBrush, e.Bounds, StringFormat.GenericDefault)
			End Using
			e.DrawFocusRectangle()
		End Sub

		Private Sub btnSaveImage_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSaveAsImage.Click
			Dim Pos As Integer = GetImagePos
			If Pos < 0 Then
				MessageBox.Show("There is no selected image to save", "Error")
				Return
			End If

			Dim ext As String = dataSet1.Tables("ImageDataTable").Rows(Pos).ItemArray(4).ToString().ToLower()
			saveImageDialog.DefaultExt = ext
			saveImageDialog.Filter = ext & " Images|*." & ext
			If saveImageDialog.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
				Return
			End If
			Dim ImgData() As Byte = CType(dataSet1.Tables("ImageDataTable").Rows(Pos).ItemArray(8), Byte())
			Using fs As New FileStream(saveImageDialog.FileName, FileMode.Create)
				fs.Write(ImgData, 0, ImgData.Length)
			End Using

		End Sub

		Private Sub btnInfo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnInfo.Click
			MessageBox.Show("FlexCelImageExplorer is a small application targeted to reduce the size on images inside Excel files." & vbLf & "On the current version you can see the image properties and extract the images to disk.")

		End Sub

		Private Sub btnStretchPreview_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnStretchPreview.Click
			If btnStretchPreview.Checked Then
				PreviewBox.SizeMode = PictureBoxSizeMode.StretchImage
			Else
				PreviewBox.SizeMode = PictureBoxSizeMode.Normal
			End If
		End Sub

	End Class

	Friend Class TImageInfo
		Friend File As FileInfo
		Friend IsValidFile As Boolean
		Friend HasCrop As Boolean
		Friend HasImages As Boolean
		Friend HasARGB As Boolean

		Public Sub New(ByVal aFile As FileInfo, ByVal aIsValidFile As Boolean, ByVal aHasCrop As Boolean, ByVal aHasImages As Boolean, ByVal aHasARGB As Boolean)
			File = aFile
			HasCrop = aHasCrop
			HasImages = aHasImages
			IsValidFile = aIsValidFile
			HasARGB = aHasARGB
		End Sub

		Public Overrides Function ToString() As String
			If Not IsValidFile Then
				Return " (*)" & File.ToString()
			End If
			Return File.ToString()
		End Function

	End Class
End Namespace
