Imports System.Collections
Imports System.ComponentModel
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports System.IO
Imports System.Reflection

Namespace GettingStarted
	''' <summary>
	''' A small example on how to create a simple file with the API.
	''' Note that you can use the APIMate tool (in Start Menu->TMS FlexCel Studio->Tools) to find out the 
	''' methods you need to call.
	''' </summary>
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles button1.Click
			Dim Xls As ExcelFile = New XlsFile(True)
			AddData(Xls)

			If cbAutoOpen.Checked Then
				AutoOpen(Xls)
			Else
				NormalOpen(Xls)
			End If
		End Sub

		Private Sub AddData(ByVal Xls As ExcelFile)
			'Create a new file. We could also open an existing file with Xls.Open
			Xls.NewFile(1, TExcelFileFormat.v2019)
			'Set some cell values.
			Xls.SetCellValue(1, 1, "Hello to the world")
			Xls.SetCellValue(2, 1, 3)
			Xls.SetCellValue(3, 1, 2.1)
			Xls.SetCellValue(4, 1, New TFormula("=Sum(A2,A3)"))

			'Load an image from disk.
			Dim AssemblyPath As String = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
			Using Img As Image = Image.FromFile(AssemblyPath & Path.DirectorySeparatorChar & ".." & Path.DirectorySeparatorChar & ".." & Path.DirectorySeparatorChar & "Test.bmp")

				'Add a new image on cell E2
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

			'You can call AddFormat as many times as you want, it will never add a format twice.
			'But if you know the format you are going to use, you can get some extra CPU cycles by
			'calling addformat once and saving the result into a variable.
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


			'Make the page print in landscape or portrait mode
			Xls.PrintLandscape = False

		End Sub


		'This is part of an advanced feature (showing the user using a file) , you do not need to use
		'this method on normal places.
		Private Function GetLockingUser(ByVal FileName As String) As String
			Try
				Dim xerr As New XlsFile()
				xerr.Open(FileName)
				Return " - File might be in use by: " & xerr.Protection.WriteAccess
			Catch
				Return String.Empty
			End Try
		End Function

		Private Sub NormalOpen(ByVal Xls As ExcelFile)
			If saveFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
				Try
					Xls.Save(saveFileDialog1.FileName)
				Catch ex As IOException 'This is not really needed, just to show the username of the user locking the file.
					Throw New IOException(ex.Message & GetLockingUser(saveFileDialog1.FileName), ex)
				End Try

				If MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
					Process.Start(saveFileDialog1.FileName)
				End If
			End If
		End Sub

		'This method will use a "trick" to create a temporary file and delete it even when it is open on Excel.
		'We will create a "template" (xlt file), and tell Excel to create a new file based on this template.
		'Then we can safely delete the xlt file, since Excel opened a copy.
		Private Sub AutoOpen(ByVal Xls As ExcelFile)
			Dim FilePath As String = Path.GetTempPath() 'GetTempFileName does not allow us to specify the "xlt" extension.
			Dim FileName As String = Path.Combine(FilePath, Guid.NewGuid().ToString() & ".xlt") 'xlt is the extension for excel templates.
			Try
				Using OutStream As New FileStream(FileName, FileMode.Create, FileAccess.Write)
					Dim Fi As New FileInfo(FileName)
					Fi.Attributes = FileAttributes.Temporary
					Xls.IsXltTemplate = True 'Make it an xlt template.
					Xls.Save(OutStream)
				End Using
				Process.Start(FileName)
			Finally
				File.Delete(FileName) 'As it is an xlt file, we can delete it even when it is open on Excel.
			End Try
		End Sub

		''' <summary>
		''' This is the method that will be called by the ASP.NET front end. It returns an array of bytes 
		''' with the report data, so the ASP.NET application can stream it to the client.
		''' </summary>
		''' <returns>The generated file as a byte array.</returns>
		Public Function WebRun() As Byte()
			Dim Xls As ExcelFile = New XlsFile(True)
			AddData(Xls)

			Using OutStream As New MemoryStream()
				Xls.Save(OutStream)
				Return OutStream.ToArray()
			End Using
		End Function

	End Class
End Namespace
