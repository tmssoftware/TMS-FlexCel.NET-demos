Imports System.IO


Namespace MainDemo
	Friend NotInheritable Class Program

		Private Sub New()
		End Sub

		''' <summary>
		''' The main entry point for the application.
		''' </summary>
		<STAThread> _
		Shared Sub Main()
			AddHandler Application.ThreadException, AddressOf Application_ThreadException
			Application.EnableVisualStyles()
			Application.Run(New DemoForm())
		End Sub

		Private Shared Sub Application_ThreadException(ByVal sender As Object, ByVal e As System.Threading.ThreadExceptionEventArgs)
			Dim ex As Exception = e.Exception
			Do While ex.InnerException IsNot Nothing
				ex = ex.InnerException
			Loop


			MessageBox.Show(ex.GetType().Name &" //" & ex.Message)
		End Sub

	End Class
End Namespace
