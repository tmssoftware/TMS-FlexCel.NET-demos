'This program is inspired on the progam by Mikhail Arkhipov
' * at http://blogs.msdn.com/mikhailarkhipov/archive/2004/08/12/213963.aspx
' * Thanks!
' 

' UPDATE: This was patched with the info on 
' * http://weblogs.asp.net/jan/archive/2004/01/28/63771.aspx
' * to make it work.
' * 
' * Thanks again...
' * 
' * UPDATE 2!
' * The NOAA broke the service again, and it has not fixed it for more than a year. 
' * I give up. We will use http://www.webservicex.net/WeatherForecast.asmx instead.
' * The code for NOAA is still there on the SetupNOAA method, just not used so you can see it (and try it if it ever starts working again)
' *
' * UPDATE 3!
' * Now WebserviceX is not working, going back to NOAA. As you can see, it isn't very trustable that a webservice will be there in the
' * future, so this demo might not work in online mode when you try it. But you can always look at in in offline mode.
' * 


Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports System.Xml
Imports System.Net
Imports System.Threading

Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Report
Imports FlexCel.Render

Imports gov.weather.www
Imports System.Globalization

Namespace ExportingWebServices
	''' <summary>
	''' An example that will read data from a webservice.
	''' </summary>
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Private Cities As New Dictionary(Of String, LatLong)(StringComparer.CurrentCultureIgnoreCase)
		Public Sub New()
			InitializeComponent()
			LoadCities()
		End Sub

		Private Sub LoadCities()
			Dim xml As New XmlDocument()
				Dim DataPath As String = Path.Combine(Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), ".."), "..") & Path.DirectorySeparatorChar
				xml.Load(Path.Combine(DataPath, "cities.xml"))
				Dim latLonList As XmlNodeList = xml.GetElementsByTagName("latLonList")
				Dim cityNameList As XmlNodeList = xml.GetElementsByTagName("cityNameList")

				If latLonList.Count <> 1 Then
					Throw New Exception("Invalid city list")
				End If
				If cityNameList.Count <> 1 Then
					Throw New Exception("Invalid city list")
				End If

				Dim lats As String = latLonList.Item(0).InnerText
				Dim cits As String = cityNameList.Item(0).InnerText

				Dim latsParsed() As String = lats.Split(" "c)
				Dim citsParsed() As String = cits.Split("|"c)

				If citsParsed.Length <> latsParsed.Length Then
					Throw New Exception("Invalid city list")
				End If

				edcity.BeginUpdate()
				Try
					For i As Integer = 0 To citsParsed.Length - 1
						Dim ll() As String = latsParsed(i).Split(","c)
						If ll.Length <> 2 Then
							Throw New Exception("Invalid city list")
						End If
						Cities.Add(citsParsed(i), New LatLong(Convert.ToDecimal(ll(0), CultureInfo.InvariantCulture), Convert.ToDecimal(ll(1), CultureInfo.InvariantCulture)))
						edcity.Items.Add(citsParsed(i))
					Next i

					edcity.Text = "New York,NY"
				Finally
					edcity.EndUpdate()
				End Try

		End Sub

		Private Sub Export(ByVal SaveDialog As SaveFileDialog, ByVal ToPdf As Boolean)
			Try
				Dim DataPath As String = Path.Combine(Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), ".."), "..") & Path.DirectorySeparatorChar

				'We will use a thread to connect, to avoid "freezing" the GUI
				Dim MyWebConnect As New WebConnectThread(reportStart, edcity.Text, DataPath, cbOffline.Checked, Cities)
				Dim WebConnect As New Thread(New ThreadStart(AddressOf MyWebConnect.SetupNOAA))
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
						xls.Open(DataPath & "Exporting Web Services.template.xls")
						reportStart.Run(xls)
						Using PdfExport As New FlexCelPdfExport(xls, True)
							PdfExport.Export(SaveDialog.FileName)
						End Using
					Else
						reportStart.Run(DataPath & "Exporting Web Services.template.xls", SaveDialog.FileName)
					End If

					If MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
						Process.Start(SaveDialog.FileName)
					End If
				End If
			Catch ex As Exception
				MessageBox.Show(ex.Message)
			End Try

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

		Private Sub edcity_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles edcity.KeyDown
			edcity.DroppedDown = False
		End Sub

	End Class

	Friend Structure LatLong
		Public Latitude As Decimal
		Public Longitude As Decimal

		Public Sub New(ByVal aLatitude As Decimal, ByVal aLongitude As Decimal)
			Latitude = aLatitude
			Longitude = aLongitude
		End Sub
	End Structure

	Friend Class WebConnectThread
		Private CityName As String
		Private DataPath As String
		Private UseOfflineData As Boolean
		Private Cities As Dictionary(Of String, LatLong)
		Private ReportStart As FlexCelReport
		Private FMainException As Exception

		Public Sub New(ByVal aReportStart As FlexCelReport, ByVal aCityName As String, ByVal aDataPath As String, ByVal aUseOfflineData As Boolean, ByVal aCities As Dictionary(Of String, LatLong))
			CityName = aCityName
			DataPath = aDataPath
			UseOfflineData = aUseOfflineData
			ReportStart = aReportStart
			Cities = aCities
		End Sub


		Public Sub SetupNOAA()
			Try
				SetupNOAA(ReportStart, CityName, DataPath, UseOfflineData, Cities)
			Catch ex As Exception
				FMainException = ex
			End Try

		End Sub

		Public Shared Sub SetupNOAA(ByVal reportStart As FlexCelReport, ByVal CityName As String, ByVal DataPath As String, ByVal UseOfflineData As Boolean, ByVal Cities As Dictionary(Of String, LatLong))
			Dim CityCoords As LatLong
			GetCity(Cities, CityName, CityCoords)
			reportStart.SetValue("Date", Date.Now)
			Dim forecasts As String
			Dim dtStart As Date = Date.Now

			If UseOfflineData Then
				Using fs As New StreamReader(Path.Combine(DataPath, "OfflineData.xml"))
					forecasts = fs.ReadToEnd()
				End Using
			Else
				Dim nd As New ndfdXML()
				forecasts = nd.NDFDgen(CityCoords.Latitude, CityCoords.Longitude, productType.glance, dtStart, dtStart.AddDays(7), unitType.m, New weatherParametersType())

#If(SAVEOFFLINEDATA) Then
				Using sw As New StreamWriter(Path.Combine(DataPath, "OfflineData.xml"))
					sw.Write(forecasts)
				End Using
#End If
			End If

			If String.IsNullOrEmpty(forecasts) Then
				Throw New Exception("Can't find the place " & CityName)
			End If

			Dim ds As New DataSet()
			'Load the data into a dataset. On this web service, we cannot just call DataSet.ReadXml as the data is not on the correct format.
			Dim xmlDoc As New XmlDocument()
				xmlDoc.LoadXml(forecasts)
				Dim HighList As XmlNodeList = xmlDoc.SelectNodes("/dwml/data/parameters/temperature[@type='maximum']/value/text()")
				Dim LowList As XmlNodeList = xmlDoc.SelectNodes("/dwml/data/parameters/temperature[@type='minimum']/value/text()")
				Dim IconList As XmlNodeList = xmlDoc.SelectNodes("/dwml/data/parameters/conditions-icon/icon-link/text()")

				Dim WeatherTable As DataTable = ds.Tables.Add("Weather")

				WeatherTable.Columns.Add("Day", GetType(Date))
				WeatherTable.Columns.Add("Low", GetType(Double))
				WeatherTable.Columns.Add("High", GetType(Double))
				WeatherTable.Columns.Add("Icon", GetType(Byte()))

				For i As Integer = 0 To Math.Min(Math.Min(HighList.Count, LowList.Count), IconList.Count) - 1
					WeatherTable.Rows.Add(New Object(){ dtStart.AddDays(i), Convert.ToDouble(LowList(i).Value), Convert.ToDouble(HighList(i).Value), LoadIcon(IconList(i).Value, UseOfflineData, DataPath)})
				Next i


			reportStart.AddTable(ds, TDisposeMode.DisposeAfterRun)
			reportStart.SetValue("Latitude", CityCoords.Latitude)
			reportStart.SetValue("Longitude", CityCoords.Longitude)
			reportStart.SetValue("Place", CityName)

		End Sub

		Private Shared Sub GetCity(ByVal Cities As Dictionary(Of String, LatLong), ByVal CityName As String, ByRef CityCoords As LatLong)
			If Not Cities.TryGetValue(CityName, CityCoords) Then
				Throw New Exception("Can't find the city " & CityName)
			End If
		End Sub

		Friend ReadOnly Property MainException() As Exception
			Get
				Return FMainException
			End Get
		End Property

		Friend Shared Function LoadIcon(ByVal url As String, ByVal useOfflineData As Boolean, ByVal dataPath As String) As Byte()
			If url Is Nothing OrElse url.Length = 0 Then
				Return Nothing 'no icon for this image.
			End If

			If useOfflineData Then
				Dim u As New Uri(url)
				Return LoadFileIcon(Path.Combine(dataPath, u.Segments(u.Segments.Length - 1)))
			Else
#If (SAVEOFFLINEDATA) Then
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
				wc.Headers.Add("user-agent", "FlexCel Webservice Example")
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
