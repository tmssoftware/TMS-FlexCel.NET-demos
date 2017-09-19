Imports System.Collections
Imports System.ComponentModel

Imports System.Threading

Imports FlexCel.Render

Namespace ExportingWebServices
	''' <summary>
	''' A dialog box to show progress. This could be done with a BackgroundWorker, it was done this way for .NET 1.1 compatibility.
	''' </summary>
	Partial Public Class ProgressDialog
		Inherits System.Windows.Forms.Form

		Private WithEvents timer1 As System.Timers.Timer

		Public Sub New()
			InitializeComponent()
		End Sub


		Private StartTime As Date
		Private RunningThread As Thread

		Private Sub timer1_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles timer1.Elapsed
			UpdateStatus()
		End Sub

		Public Sub ShowProgress(ByVal aRunningThread As Thread)
			RunningThread = aRunningThread

			If Not RunningThread.IsAlive Then
				DialogResult = System.Windows.Forms.DialogResult.OK
				Return
			End If
			timer1.Enabled = True
			StartTime = Date.Now
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
		End Sub

		Private Sub ProgressDialog_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
			timer1.Enabled = False
		End Sub

	End Class
End Namespace
