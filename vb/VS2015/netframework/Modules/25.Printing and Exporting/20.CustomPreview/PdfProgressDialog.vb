Imports System.ComponentModel

Imports System.Threading

Imports FlexCel.Render

Namespace CustomPreview
	''' <summary>
	''' Shows progress as we are exporting to pdf.
	''' </summary>
	Partial Public Class PdfProgressDialog
		Inherits System.Windows.Forms.Form

		Private WithEvents timer1 As System.Timers.Timer

		Public Sub New()
			InitializeComponent()
		End Sub


		Private StartTime As Date
		Private RunningThread As Thread
		Private PdfExport As FlexCelPdfExport

		Private Sub timer1_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles timer1.Elapsed
			UpdateStatus()
		End Sub

		Public Sub ShowProgress(ByVal aRunningThread As Thread, ByVal aPdfExport As FlexCelPdfExport)
			RunningThread = aRunningThread

			If Not RunningThread.IsAlive Then
				DialogResult = System.Windows.Forms.DialogResult.OK
				Return
			End If
			timer1.Enabled = True
			StartTime = Date.Now
			PdfExport = aPdfExport
			ShowDialog()
		End Sub

		Private Sub UpdateStatus()
			Dim ts As TimeSpan = Date.Now.Subtract(StartTime)
			Dim hours As String
			If ts.Hours = 0 Then
				hours = ""
			Else
				hours = ts.Hours.ToString("00") & ":"
			End If
			statusBarPanelTime.Text = hours & ts.Minutes.ToString("00") & ":" & ts.Seconds.ToString("00")

			If Not RunningThread.IsAlive Then
				DialogResult = System.Windows.Forms.DialogResult.OK
			End If

			If PdfExport.Progress.TotalPage > 0 Then
				labelPages.Text = String.Format("Generating Page {0} of {1}", PdfExport.Progress.Page, PdfExport.Progress.TotalPage)
			End If
		End Sub

		Private Sub PdfProgressDialog_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
			timer1.Enabled = False
		End Sub


	End Class
End Namespace
