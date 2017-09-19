Imports System.Net.Sockets
Imports System.IO
Imports System.Threading

Imports System.Globalization
Imports System.Text

Namespace ExportHTML
	''' <summary>
	''' IMPORTANT: This is a really simple SMTP implementation, and it is not indented to be used in production.
	''' It is just mean to be use on this simple demo, as we cannot use third party solutions here.
	''' You are advised to get a third party solution if you want to implement this in your application.
	''' This code is based directly in RFC 2821: http://www.faqs.org/rfcs/rfc2821.html
	''' </summary>
	Public Class SimpleMailer
		#Region "privates"
		Private FHostName As String
		Private FPort As Integer

		Private FToAddress As String
		Private FFromAddress As String
		Private FSubject As String
		#End Region

		#Region "Public properties"
		Public Property HostName() As String
			Get
				Return FHostName
			End Get
			Set(ByVal value As String)
				FHostName = value
			End Set
		End Property
		Public Property Port() As Integer
			Get
				Return FPort
			End Get
			Set(ByVal value As Integer)
				FPort = value
			End Set
		End Property
		Public Property ToAddress() As String
			Get
				Return FToAddress
			End Get
			Set(ByVal value As String)
				FToAddress = value
			End Set
		End Property
		Public Property FromAddress() As String
			Get
				Return FFromAddress
			End Get
			Set(ByVal value As String)
				FFromAddress = value
			End Set
		End Property
		Public Property Subject() As String
			Get
				Return FSubject
			End Get
			Set(ByVal value As String)
				FSubject = value
			End Set
		End Property
		#End Region

		Public Sub New()
			Port = 25
		End Sub

		Public Sub SendMail(ByVal MessageData() As Byte)
			Using Tcp As New TcpClient(HostName, Port)
				Using Channel As NetworkStream = Tcp.GetStream()
					Try
						WaitForAnswer(Channel, 220)
						WriteCommand(Channel, "EHLO " & HostName)
						WaitForAnswer(Channel, 250)

						WriteCommand(Channel, "MAIL FROM:<" & FromAddress & ">")
						WaitForAnswer(Channel, 250)
						WriteCommand(Channel, "RCPT TO:<" & ToAddress & ">")
						WaitForAnswer(Channel, 250)

						WriteCommand(Channel, "DATA")
						WaitForAnswer(Channel, 354)
						SendMessage(Channel, MessageData)
						WaitForAnswer(Channel, 250)

					Catch e1 As SimpleMailerException
						'Quit must always be performed.
						WriteCommand(Channel, "QUIT")
						WaitForAnswer(Channel, 221)
						Return
					End Try

					WriteCommand(Channel, "QUIT")
					WaitForAnswer(Channel, 221)

				End Using
			End Using
		End Sub

		Public Sub WaitForAnswer(ByVal Channel As NetworkStream, ByVal AnswerCode As Integer)
			Dim ReadBuffer(1023) As Byte
			Dim Message As String = ""
			Dim numberOfBytesRead As Integer = 0

			' Incoming message may be larger than the buffer size.
			Do
				numberOfBytesRead = Channel.Read(ReadBuffer, 0, ReadBuffer.Length)
				Message &= Encoding.ASCII.GetString(ReadBuffer, 0, numberOfBytesRead)
			Loop While Channel.DataAvailable

			If Message.StartsWith(AnswerCode.ToString(CultureInfo.InvariantCulture)) Then
				Return
			End If

			Throw New SimpleMailerException("Error sending email. Answer from the server: " & Message)
		End Sub

		Public Sub WriteCommand(ByVal Channel As NetworkStream, ByVal Command As String)
			Dim WriteBuffer() As Byte = Encoding.ASCII.GetBytes(Command & vbCrLf) 'do not use Environment.NewLine here, since this is invariant and defined in the RFC
			Channel.Write(WriteBuffer, 0, WriteBuffer.Length)
		End Sub

		Public Sub SendMessage(ByVal Channel As NetworkStream, ByVal MessageData() As Byte)
			WriteCommand(Channel, "From: " & FromAddress)
			WriteCommand(Channel, "To: " & ToAddress)
			WriteCommand(Channel, "Subject: " & Subject)
			WriteCommand(Channel, "Date: " & Date.Now.ToUniversalTime().ToString("r", CultureInfo.InvariantCulture))

			WriteEscapedData(Channel, MessageData)
			WriteCommand(Channel, vbCrLf & ".") 'do not use Environment.NewLine here, since this is invariant and defined in the RFC
		End Sub

		''' <summary>
		''' In order to send the data trough the channel, we need to detect any dot at the start of a line and replace it by ".."
		''' If not, it might be interpreted as an EOF sign. We can easily loop in the MessageData array, since it only contains ASCII characters, so there are no
		''' unicode colation issues.
		''' </summary>
		''' <param name="Channel"></param>
		''' <param name="MessageData"></param>
		Public Sub WriteEscapedData(ByVal Channel As NetworkStream, ByVal MessageData() As Byte)
			If MessageData.Length > 0 Then
				If MessageData(0) = AscW("."c) Then
					Channel.WriteByte(AscW("."c))
				End If
				Channel.WriteByte(MessageData(0))
			End If

			For i As Integer = 1 To MessageData.Length - 1
					If MessageData(i - 1) = AscW(ControlChars.Lf) AndAlso MessageData(i) = AscW("."c) Then
					Channel.WriteByte(AscW("."c))
				End If
				Channel.WriteByte(MessageData(i))
			Next i
		End Sub

	End Class

	Public Class SimpleMailerException
		Inherits IOException

		Public Sub New(ByVal Message As String)
			MyBase.New(Message)

		End Sub
	End Class
End Namespace
