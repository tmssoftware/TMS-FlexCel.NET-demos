Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports System.Xml
Imports System.Net
Imports System.Threading
Imports System.Globalization

Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Report
Imports FlexCel.Render

Namespace HTML
	''' <summary>
	''' Shows the limited HTML support.
	''' </summary>
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub Export(ByVal SaveDialog As SaveFileDialog, ByVal ToPdf As Boolean)
			Using reportStart As New FlexCelReport(True)

				If cbOffline.Checked AndAlso edCity.Text <> "london" Then
					MessageBox.Show("Offline mode is selected, so we will show the data of london. The actual city you wrote will not be used unless you select online mode.", "Warning")
				End If
				Try
					Dim DataPath As String = Path.Combine(Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), ".."), "..") & Path.DirectorySeparatorChar
					Dim OfflineDataPath As String = Path.Combine(DataPath, "OfflineData") & Path.DirectorySeparatorChar

					'We will use a thread to connect, to avoid "freezing" the GUI
					Dim MyWebConnect As New WebConnectThread(reportStart, edCity.Text, OfflineDataPath, cbOffline.Checked)
					Dim WebConnect As New Thread(New ThreadStart(AddressOf MyWebConnect.LoadData))
					WebConnect.Start()
					Using Pg As New ProgressDialog()
						Pg.ShowProgress(WebConnect)
						If MyWebConnect IsNot Nothing AndAlso MyWebConnect.MainException IsNot Nothing Then
							Throw MyWebConnect.MainException
						End If
					End Using


					If SaveDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
						If ToPdf Then
							Dim xls As New XlsFile()
							xls.Open(DataPath & "HTML.template.xls")
							reportStart.Run(xls)
							Using PdfExport As New FlexCelPdfExport(xls, True)
								PdfExport.Export(SaveDialog.FileName)
							End Using
						Else
							reportStart.Run(DataPath & "HTML.template.xls", SaveDialog.FileName)
						End If

						If MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
							Process.Start(SaveDialog.FileName)
						End If
					End If
				Catch ex As Exception
					MessageBox.Show(ex.Message)
				End Try
			End Using
		End Sub

		Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
			Close()
		End Sub

		Private Sub btnExportPdf_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExportPdf.Click
			Export(saveFileDialogPdf, True)
		End Sub

		Private Sub btnExportXls_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExportXls.Click
			Export(saveFileDialogXls, False)
		End Sub

		Private Sub linkLabel1_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles linkLabel1.LinkClicked
			System.Diagnostics.Process.Start((TryCast(sender, LinkLabel)).Text)
		End Sub
	End Class

	Friend Class WebConnectThread
		Private CityName As String
		Private DataPath As String
		Private UseOfflineData As Boolean
		Private ReportStart As FlexCelReport
		Private FMainException As Exception

		Public Sub New(ByVal aReportStart As FlexCelReport, ByVal aCityName As String, ByVal aDataPath As String, ByVal aUseOfflineData As Boolean)
			CityName = aCityName
			DataPath = aDataPath
			UseOfflineData = aUseOfflineData
			ReportStart = aReportStart
		End Sub


		''' <summary>
		''' This is the method we will call form a thread. It catches any internal exception.
		''' </summary>
		Public Sub LoadData()
			Try
				LoadData(ReportStart, CityName, DataPath, UseOfflineData)
			Catch ex As Exception
				FMainException = ex
			End Try

		End Sub

		Public Shared Sub LoadData(ByVal reportStart As FlexCelReport, ByVal CityName As String, ByVal DataPath As String, ByVal UseOfflineData As Boolean)
			reportStart.SetValue("Date", Date.Now)
			Dim ds As New DataSet()
			ds.Locale = CultureInfo.InvariantCulture
			ds.EnforceConstraints = False
			ds.ReadXmlSchema(Path.Combine(DataPath, "TripSearchResponse.xsd"))
			ds.Tables("Result").Columns.Add("ImageData", GetType(Byte())) 'Add a column for the actual images.
			If UseOfflineData Then
				ds.ReadXml(Path.Combine(DataPath, "OfflineData.xml"))
			Else
				' Create the web request  
				Dim url As String = String.Format("http://travel.yahooapis.com/TripService/V1.1/tripSearch?appid=YahooDemo&query={0}&results=20", CityName)
				Dim uri As New UriBuilder(url)
				Dim request As HttpWebRequest = TryCast(WebRequest.Create(uri.Uri.AbsoluteUri), HttpWebRequest)

				' Get response  
				Using response As HttpWebResponse = TryCast(request.GetResponse(), HttpWebResponse)
					' Load data into a dataset  
					ds.ReadXml(response.GetResponseStream())
				End Using
			End If

			If ds.Tables("ResultSet").Rows.Count <= 0 Then
				Throw New Exception("Error loading the data.")
			End If
			If Convert.ToInt32(ds.Tables("ResultSet").Rows(0)("totalResultsReturned")) <= 0 Then
				Throw New Exception("There are no travel plans for this location")
			End If

			LoadImageData(ds, UseOfflineData, DataPath)

			' Uncomment this code to create an offline image of the data.
#If (CreateOffline) Then
			ds.WriteXml(Path.Combine(DataPath, "OfflineData.xml"))
#End If

			reportStart.AddTable(ds)
		End Sub

		Friend ReadOnly Property MainException() As Exception
			Get
				Return FMainException
			End Get
		End Property

		Private Shared Sub LoadImageData(ByVal ds As DataSet, ByVal UseOfflineData As Boolean, ByVal DataPath As String)
			Dim Images As DataTable = ds.Tables("Image")
			Images.PrimaryKey = New DataColumn() { Images.Columns("Result_Id") }
			For Each dr As DataRow In ds.Tables("Result").Rows
				Dim ImageRow As DataRow = Images.Rows.Find(dr("Result_Id"))

				If ImageRow Is Nothing Then
					Continue For
				End If
				Dim url As String = Convert.ToString(ImageRow("Url"))
				If url IsNot Nothing AndAlso url.Length > 0 Then
					dr("ImageData") = LoadIcon(url, UseOfflineData, DataPath)
				End If
			Next dr

		End Sub

		Friend Shared Function LoadIcon(ByVal url As String, ByVal useOfflineData As Boolean, ByVal dataPath As String) As Byte()
			If useOfflineData Then
				Dim u As New Uri(url)
				Return LoadFileIcon(Path.Combine(dataPath, u.Segments(u.Segments.Length - 1)))
			Else
				' Uncomment this code to create an offline image of the data. 
#If (CreateOffline) Then
				Dim u As New Uri(url)
				Dim IconData() As Byte = LoadWebIcon(url)
				Using fs As New FileStream(Path.Combine(dataPath, u.Segments(u.Segments.Length - 1)), FileMode.Create)
					fs.Write(IconData, 0, IconData.Length)
				End Using
#End If


				Return LoadWebIcon(url)

			End If
		End Function

		''' <summary>
		''' On a real implementation this should be cached.
		''' </summary>
		''' <param name="url"></param>
		''' <returns></returns>
		Friend Shared Function LoadWebIcon(ByVal url As String) As Byte()
			Using wc As New WebClient()
				Return wc.DownloadData(url)
			End Using
		End Function

		Private Shared Function LoadFileIcon(ByVal filename As String) As Byte()
			Using fs As New FileStream(filename, FileMode.Open)
				Dim Result(CInt(fs.Length) - 1) As Byte
				fs.Read(Result, 0, Result.Length)
				Return Result
			End Using
		End Function


	End Class

End Namespace
