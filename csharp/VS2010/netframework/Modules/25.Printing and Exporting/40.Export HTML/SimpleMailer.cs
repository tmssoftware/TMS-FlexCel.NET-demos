using System;
using System.Net.Sockets;
using System.IO;
using System.Threading;

using System.Globalization;
using System.Text;

namespace ExportHTML
{
    /// <summary>
    /// IMPORTANT: This is a really simple SMTP implementation, and it is not indented to be used in production.
    /// It is just mean to be use on this simple demo, as we cannot use third party solutions here.
    /// You are advised to get a third party solution if you want to implement this in your application.
    /// This code is based directly in RFC 2821: http://www.faqs.org/rfcs/rfc2821.html
    /// </summary>
    public class SimpleMailer
    {
        #region privates
        private string FHostName;
        private int FPort;

        private string FToAddress;
        private string FFromAddress;
        private string FSubject;
        #endregion

        #region Public properties
        public string HostName { get { return FHostName; } set { FHostName = value; } }
        public int Port { get { return FPort; } set { FPort = value; } }
        public string ToAddress { get { return FToAddress; } set { FToAddress = value; } }
        public string FromAddress { get { return FFromAddress; } set { FFromAddress = value; } }
        public string Subject { get { return FSubject; } set { FSubject = value; } }
        #endregion

        public SimpleMailer()
        {
            Port = 25;
        }

        public void SendMail(byte[] MessageData)
        {
            using (TcpClient Tcp = new TcpClient(HostName, Port))
            {
                using (NetworkStream Channel = Tcp.GetStream())
                {
                    try
                    {
                        WaitForAnswer(Channel, 220);
                        WriteCommand(Channel, "EHLO " + HostName);
                        WaitForAnswer(Channel, 250);

                        WriteCommand(Channel, "MAIL FROM:<" + FromAddress + ">");
                        WaitForAnswer(Channel, 250);
                        WriteCommand(Channel, "RCPT TO:<" + ToAddress + ">");
                        WaitForAnswer(Channel, 250);

                        WriteCommand(Channel, "DATA");
                        WaitForAnswer(Channel, 354);
                        SendMessage(Channel, MessageData);
                        WaitForAnswer(Channel, 250);
                    }

                    catch (SimpleMailerException)
                    {
                        //Quit must always be performed.
                        WriteCommand(Channel, "QUIT");
                        WaitForAnswer(Channel, 221);
                        return;
                    }

                    WriteCommand(Channel, "QUIT");
                    WaitForAnswer(Channel, 221);

                }
            }
        }

        public void WaitForAnswer(NetworkStream Channel, int AnswerCode)
        {
            byte[] ReadBuffer = new byte[1024];
            String Message = "";
            int numberOfBytesRead = 0;

            // Incoming message may be larger than the buffer size.
            do
            {
                numberOfBytesRead = Channel.Read(ReadBuffer, 0, ReadBuffer.Length);
                Message += Encoding.ASCII.GetString(ReadBuffer, 0, numberOfBytesRead);
            }
            while (Channel.DataAvailable);

            if (Message.StartsWith(AnswerCode.ToString(CultureInfo.InvariantCulture))) return;

            throw new SimpleMailerException("Error sending email. Answer from the server: " + Message);
        }

        public void WriteCommand(NetworkStream Channel, string Command)
        {
            byte[] WriteBuffer = Encoding.ASCII.GetBytes(Command + "\r\n"); //do not use Environment.NewLine here, since this is invariant and defined in the RFC
            Channel.Write(WriteBuffer, 0, WriteBuffer.Length);
        }

        public void SendMessage(NetworkStream Channel, byte[] MessageData)
        {
            WriteCommand(Channel, "From: " + FromAddress);
            WriteCommand(Channel, "To: " + ToAddress);
            WriteCommand(Channel, "Subject: " + Subject);
            WriteCommand(Channel, "Date: " + DateTime.Now.ToUniversalTime().ToString("r", CultureInfo.InvariantCulture));

            WriteEscapedData(Channel, MessageData);
            WriteCommand(Channel, "\r\n."); //do not use Environment.NewLine here, since this is invariant and defined in the RFC
        }

        /// <summary>
        /// In order to send the data trough the channel, we need to detect any dot at the start of a line and replace it by ".."
        /// If not, it might be interpreted as an EOF sign. We can easily loop in the MessageData array, since it only contains ASCII characters, so there are no
        /// unicode colation issues.
        /// </summary>
        /// <param name="Channel"></param>
        /// <param name="MessageData"></param>
        public void WriteEscapedData(NetworkStream Channel, byte[] MessageData)
        {
            if (MessageData.Length > 0)
            {
                if (MessageData[0] == '.') Channel.WriteByte((byte)'.');
                Channel.WriteByte(MessageData[0]);
            }

            for (int i = 1; i < MessageData.Length; i++)
            {
                if (MessageData[i - 1] == '\n' && MessageData[i] == '.')
                {
                    Channel.WriteByte((byte)'.');
                }
                Channel.WriteByte(MessageData[i]);
            }
        }

    }

    public class SimpleMailerException: IOException
    {
        public SimpleMailerException(string Message) : base(Message)
        {

        }
    }
}
