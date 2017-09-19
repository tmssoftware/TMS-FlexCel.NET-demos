Imports System.Collections
Imports System.ComponentModel

Imports FlexCel.Core
Imports System.Text

Namespace ExportHTML
	''' <summary>
	''' Used for mailing.
	''' </summary>
	Partial Public Class Mailform
		Inherits System.Windows.Forms.Form

		Private OriginalTo As String
		Private OriginalFrom As String
		Private OriginalServer As String

		Public Sub New()
			InitializeComponent()

			OriginalTo = edTo.Text
			OriginalFrom = edFrom.Text
			OriginalServer = edOutServer.Text
		End Sub


		Public MainForm As mainForm

		Private Function ValidateFields() As Boolean
			If OriginalTo = edTo.Text Then
				MessageBox.Show("Please change the 'TO' field to the user you want to send the email")
				edTo.Focus()
				Return False
			End If

			If OriginalFrom = edFrom.Text Then
				MessageBox.Show("Please change the 'From' field to the user you are using to send the email")
				edFrom.Focus()
				Return False
			End If

			If OriginalServer = edOutServer.Text Then
				MessageBox.Show("Please change the 'Outgoing Mail Server' field to the pop3 server you will use to send the email.")
				edOutServer.Focus()
				Return False
			End If

			Return True

		End Function

		Private Sub btnEmail_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEmail.Click
			If Not ValidateFields() Then
				Return
			End If

			If MessageBox.Show("Now we will try to send the email using the server '" & edOutServer.Text & "'" & vbLf & "Note that this is a very simple implementation, and it will not work if the SMTP server needs to login." & vbLf & "For this to work, you need a mail server that authenticates when reading the email, and then login into your normal account with your normal mail reader." & vbLf & vbLf & "If you need to authenticate in order to send mail, you will need to modify this code.", "Information", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) <> System.Windows.Forms.DialogResult.OK Then
				Return
			End If



			Dim Mailer As New SimpleMailer()

			Mailer.FromAddress = edFrom.Text
			Mailer.ToAddress = edTo.Text
			Mailer.Subject = edSubject.Text
			Mailer.HostName = edOutServer.Text
			Mailer.Port = 25

			Try
				Mailer.SendMail(MainForm.GenerateMHTML())
			Catch ex As Exception
				MessageBox.Show("Error trying to send the message: " & ex.Message)
				Return
			End Try

			MessageBox.Show("Message has been sent. Please verify your JUNK folder or any filters, since it might be filtered as spam")
			Close()
		End Sub

		Private Sub edFrom_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles edFrom.Leave
			If OriginalTo = edTo.Text AndAlso OriginalFrom <> edFrom.Text Then
				edTo.Text = edFrom.Text
			End If
			FillServer()
		End Sub

		Private Sub FillServer()
			If OriginalServer = edOutServer.Text AndAlso OriginalFrom <> edFrom.Text Then
				Dim AtPos As Integer = edFrom.Text.IndexOf("@")
				If AtPos > 0 Then
					Dim Server As String = edFrom.Text.Substring(AtPos + 1)
					edOutServer.Text = "mail." & Server
				End If
			End If
		End Sub
		Private Sub edTo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles edTo.Leave
			If OriginalFrom = edFrom.Text AndAlso OriginalTo <> edTo.Text Then
				edFrom.Text = edTo.Text
			End If
			FillServer()

		End Sub

	End Class
End Namespace
