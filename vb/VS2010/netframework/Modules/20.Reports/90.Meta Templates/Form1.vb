Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports System.Data.OleDb
Imports System.Threading
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Report

Imports System.Xml


Namespace MetaTemplates
	''' <summary>
	''' Templates that self-modify themselves before running.
	''' </summary>
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Public Structure FeedData
			Public Name As String
			Public Url As String
			Public Logo As String

			Public Sub New(ByVal aName As String, ByVal aUrl As String, ByVal aLogo As String)
				Name = aName
				Url = aUrl
				Logo = aLogo
			End Sub

			Public Overrides Function ToString() As String
				Return Name
			End Function
		End Structure

		Private Feeds() As FeedData = { _
			New FeedData("TMS", "http://www.tmssoftware.com/rss/tms.xml", "tms.gif"), _
			New FeedData("MSDN","https://sxpdata.microsoft.com/feeds/3.0/msdntn/MSDNMagazine_enus", "msdn.jpg"), _
			New FeedData("SLASHDOT", "http://rss.slashdot.org/Slashdot/slashdot", "slashdot.gif") _
		}


		Private Sub button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles button2.Click
			Close()
		End Sub

		Private ReadOnly Property DataPath() As String
			Get
				Return Path.Combine(Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), ".."), "..") & Path.DirectorySeparatorChar
			End Get
		End Property

		Private Sub Export(ByVal data As DataSet)
			Using Report As New FlexCelReport(True)
				Report.AddTable(data)
				Report.SetValue("FeedName", CType(cbFeeds.SelectedValue, FeedData).Name)
				Report.SetValue("FeedUrl", CType(cbFeeds.SelectedValue, FeedData).Url)
				Report.SetValue("ShowCount", cbShowFeedCount.Checked)

				Using fs As New FileStream(Path.Combine(Path.Combine(DataPath, "logos"), CType(cbFeeds.SelectedValue, FeedData).Logo), FileMode.Open)
					Dim b(CInt(fs.Length) - 1) As Byte
					fs.Read(b, 0, b.Length)
					Report.SetValue("Logo", b)
				End Using
				Report.Run(DataPath & "Meta Templates.template.xls", saveFileDialog1.FileName)
			End Using

		End Sub

		Private Sub btnExportExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExportExcel.Click

			Using data As New DataSet()
				Dim LocalData As String = Path.Combine(Path.Combine(DataPath, "data"), CType(cbFeeds.SelectedValue, FeedData).Name & ".xml")

				If cbOffline.Checked Then
					data.ReadXml(LocalData)
				Else
					'In a real world example, this should be done on a thread, as it is done in the HTML example.
					'To keep things simple here ,we will just "freeze" the gui while downloading the data, without
					'providing feedback to the user.
					Dim FeedReader As New XmlTextReader(CType(cbFeeds.SelectedValue, FeedData).Url)
					data.ReadXml(FeedReader)
				End If

#If (SaveForOffline) Then
				data.WriteXml(LocalData)
#End If


				If saveFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
					Export(data)

					If MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
						Process.Start(saveFileDialog1.FileName)
					End If
				End If
			End Using
		End Sub

		Private Sub mainForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
			cbFeeds.DataSource = Feeds
			cbFeeds.SelectedIndex = 0
		End Sub

	End Class
End Namespace
