Namespace MainDemo
	Friend NotInheritable Class Builder

		Private Sub New()
		End Sub

		Friend Shared Sub Build(ByVal Proj As String)
#If(FRAMEWORK40) Then
			Dim prj = New Microsoft.Build.Evaluation.Project(Proj)
			Try
#Else
			Dim prj As New Microsoft.Build.BuildEngine.Project()
			prj.Load(Proj)
#End If

				prj.SetProperty("DefaultTargets", "Build")
				prj.SetProperty("Configuration", "Debug")
				If Not prj.Build() Then
					Throw New Exception("Error building project: " & Proj & vbLf)
				End If
#If(FRAMEWORK40) Then
			Finally
				Microsoft.Build.Evaluation.ProjectCollection.GlobalProjectCollection.UnloadProject(prj)
			End Try
#End If
		End Sub
	End Class
End Namespace

