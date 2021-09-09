Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports System.Threading
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Report


Namespace GettingStartedReports
	''' <summary>
	''' Simple report
	''' </summary>
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnGo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnGo.Click
			'Note that we are using a FlexCelReport component in a form here. We could also create the FlexCelReport component dynamically.

			If cbAutoOpen.Checked Then
				AutoOpenRun()
			Else
				NormalRun()
			End If
		End Sub

		Private Sub Setup(ByVal UserName As String, ByVal UserUrl As String, ByVal DataPath As String)
			'Set report variables, including an image.

			reportStart.SetValue("Date", Date.Now)
			reportStart.SetValue("Name", UserName)
			reportStart.SetValue("TwoLines", "First line" & Environment.NewLine & "Second Line")
			reportStart.SetValue("Empty", Nothing)
			reportStart.SetValue("LinkPage", UserUrl)
			reportStart.SetValue("Img", File.ReadAllBytes(Path.Combine(DataPath, "img.png")))
		End Sub


		Private Sub NormalRun()
			Dim DataPath As String = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
			DataPath = Path.Combine(DataPath, "..")
			DataPath = Path.Combine(DataPath, "..")
			Setup(edName.Text, edUrl.Text, DataPath)

			If saveFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
				'FlexCel isn't a conversion tool. While it does a good job converting a lot of stuff
				'between xls and xlsx, for best results we will use an xlsx template if the user choose xlsx and xls if the user choose xls.
				reportStart.Run(Path.Combine(DataPath, "Getting Started Reports.template" & Path.GetExtension(saveFileDialog1.FileName)), saveFileDialog1.FileName)

				If MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
					Process.Start(saveFileDialog1.FileName)
				End If
			End If
		End Sub



        Private Sub AutoOpenRun()
            Dim DataPath As String = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
            DataPath = Path.Combine(DataPath, "..")
            DataPath = Path.Combine(DataPath, "..")
            Setup(edName.Text, edUrl.Text, DataPath)

            Dim Xls As New XlsFile()
            Xls.Open(Path.Combine(DataPath, "Getting Started Reports.template.xlsx"))
            reportStart.Run(Xls)

            Dim FilePath As String = Path.GetTempPath() 'GetTempFileName does not allow us to specify the "xltx" extension.
            Dim FileName As String = Path.Combine(FilePath, Guid.NewGuid().ToString() & ".xltx") 'xltx is the extension for excel templates.
            Try
                Using OutStream As New FileStream(FileName, FileMode.Create, FileAccess.ReadWrite)
                    Xls.IsXltTemplate = True 'Make it an xltx template.
                    Xls.Save(OutStream)
                End Using
                Process.Start(FileName)
            Finally
                'For .Net 4 and newer you can use Task.Run here. See https://doc.tmssoftware.com/flexcel/net/tips/automatically-open-generated-excel-files.html
                Dim t As New Thread(AddressOf RemoveTempAfterUse)
                t.Start(FileName)
            End Try
		End Sub

        Private Sub RemoveTempAfterUse(ByVal FileName As Object)
            'As it is an xltx file, we can delete it even when it is open on Excel. - wait for 30 secs to give Excel time to start.
            Thread.Sleep(30000)
            File.Delete(CType(FileName, String))
        End Sub

        ''' <summary>
        ''' This is the method that will be called by the ASP.NET front end. It returns an array of bytes 
        ''' with the report data, so the ASP.NET application can stream it to the client.
        ''' </summary>
        ''' <param name="UserName"></param>
        ''' <param name="UserUrl"></param>
        ''' <returns>The generated file as a byte array.</returns>
        Public Function WebRun(ByVal UserName As String, ByVal UserUrl As String) As Byte()
            Dim DataPath As String = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
            DataPath = Path.Combine(DataPath, "..")
            DataPath = Path.Combine(DataPath, "..")
            Setup(UserName, UserUrl, DataPath)

            Using OutStream As New MemoryStream()
                Using InStream As New FileStream(Path.Combine(DataPath, "Getting Started Reports.template.xls"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                    reportStart.Run(InStream, OutStream)
                    Return OutStream.ToArray()
                End Using
            End Using
        End Function


        Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
			Close()
		End Sub

	End Class

End Namespace
