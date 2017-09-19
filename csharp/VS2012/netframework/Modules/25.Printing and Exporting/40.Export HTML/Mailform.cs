using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

using FlexCel.Core;
using System.Text;

namespace ExportHTML
{
    /// <summary>
    /// Used for mailing.
    /// </summary>
    public partial class Mailform: System.Windows.Forms.Form
    {
        private string OriginalTo;
        private string OriginalFrom;
        private string OriginalServer;

        public Mailform()
        {
            InitializeComponent();

            OriginalTo = edTo.Text;
            OriginalFrom = edFrom.Text;
            OriginalServer = edOutServer.Text;
        }


        public mainForm MainForm;

        private bool ValidateFields()
        {
            if (OriginalTo == edTo.Text)
            {
                MessageBox.Show("Please change the 'TO' field to the user you want to send the email");
                edTo.Focus();
                return false;
            }

            if (OriginalFrom == edFrom.Text)
            {
                MessageBox.Show("Please change the 'From' field to the user you are using to send the email");
                edFrom.Focus();
                return false;
            }

            if (OriginalServer == edOutServer.Text)
            {
                MessageBox.Show("Please change the 'Outgoing Mail Server' field to the pop3 server you will use to send the email.");
                edOutServer.Focus();
                return false;
            }

            return true;

        }

        private void btnEmail_Click(object sender, System.EventArgs e)
        {
            if (!ValidateFields()) return;

            if (MessageBox.Show("Now we will try to send the email using the server '" + edOutServer.Text + "'\n" +
                "Note that this is a very simple implementation, and it will not work if the SMTP server needs to login.\n" +
                "For this to work, you need a mail server that authenticates when reading the email, and then login into your normal account with your normal mail reader.\n\n" +
                "If you need to authenticate in order to send mail, you will need to modify this code.",
                "Information", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) != DialogResult.OK) return;



            SimpleMailer Mailer = new SimpleMailer();

            Mailer.FromAddress = edFrom.Text;
            Mailer.ToAddress = edTo.Text;
            Mailer.Subject = edSubject.Text;
            Mailer.HostName = edOutServer.Text;
            Mailer.Port = 25;

            try
            {
                Mailer.SendMail(MainForm.GenerateMHTML());
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error trying to send the message: " + ex.Message);
                return;
            }

            MessageBox.Show("Message has been sent. Please verify your JUNK folder or any filters, since it might be filtered as spam");
            Close();
        }

        private void edFrom_Leave(object sender, System.EventArgs e)
        {
            if (OriginalTo == edTo.Text && OriginalFrom != edFrom.Text)
            {
                edTo.Text = edFrom.Text;
            }
            FillServer();
        }

        private void FillServer()
        {
            if (OriginalServer == edOutServer.Text && OriginalFrom != edFrom.Text)
            {
                int AtPos = edFrom.Text.IndexOf("@");
                if (AtPos > 0)
                {
                    string Server = edFrom.Text.Substring(AtPos + 1);
                    edOutServer.Text = "mail." + Server;
                }
            }
        }
        private void edTo_Leave(object sender, System.EventArgs e)
        {
            if (OriginalFrom == edFrom.Text && OriginalTo != edTo.Text)
            {
                edFrom.Text = edTo.Text;
            }
            FillServer();

        }

    }
}
