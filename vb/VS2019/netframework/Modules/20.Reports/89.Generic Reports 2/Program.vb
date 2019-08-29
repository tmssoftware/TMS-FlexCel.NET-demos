Imports System.Threading

Namespace GenericReports2
	Friend NotInheritable Class Program

		Private Sub New()
		End Sub

		''' <summary>
		''' The main entry point for the application.
		''' </summary>
		<STAThread> _
		Shared Sub Main()
			Dim handler As New ThreadExceptionHandler()

			AddHandler Application.ThreadException, AddressOf handler.Application_ThreadException

			Application.EnableVisualStyles()
			Application.SetCompatibleTextRenderingDefault(False)
			Application.Run(New mainForm())
		End Sub
	End Class

	Friend Class ThreadExceptionHandler
		Public Sub Application_ThreadException(ByVal sender As Object, ByVal e As ThreadExceptionEventArgs)
			Try
				Dim result As DialogResult = ShowThreadExceptionDialog(e.Exception)

				If result = DialogResult.Abort Then
					Application.Exit()
				End If
			Catch
				' Fatal error, terminate program
				Try
					MessageBox.Show("Fatal Error", "Fatal Error", MessageBoxButtons.OK, MessageBoxIcon.Stop)
				Finally
					Application.Exit()
				End Try
			End Try
		End Sub

		''' 
		''' Creates and displays the error message.
		''' 
		Private Function ShowThreadExceptionDialog(ByVal ex As Exception) As DialogResult
			Dim errorMessage As String = "Unhandled Exception:" & vbLf & vbLf & ex.Message & vbLf & vbLf & ex.GetType().ToString() & vbLf & vbLf & "Stack Trace:" & vbLf & ex.StackTrace

			Return MessageBox.Show(errorMessage, "Application Error", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Stop)
		End Function
	End Class

End Namespace
