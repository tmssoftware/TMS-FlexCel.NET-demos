using System;
using System.Windows.Forms;
using System.Threading;

namespace GenericReports2
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            ThreadExceptionHandler handler = new ThreadExceptionHandler();

            Application.ThreadException += new ThreadExceptionEventHandler(handler.Application_ThreadException);

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new mainForm());
        }
    }

    internal class ThreadExceptionHandler
    {
        public void Application_ThreadException(object sender, ThreadExceptionEventArgs e)
        {
            try
            {
                DialogResult result = ShowThreadExceptionDialog(
                    e.Exception);

                if (result == DialogResult.Abort)
                    Application.Exit();
            }
            catch
            {
                // Fatal error, terminate program
                try
                {
                    MessageBox.Show("Fatal Error",
                        "Fatal Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Stop);
                }
                finally
                {
                    Application.Exit();
                }
            }
        }

        /// 
        /// Creates and displays the error message.
        /// 
        private DialogResult ShowThreadExceptionDialog(Exception ex)
        {
            string errorMessage =
                "Unhandled Exception:\n\n" +
                ex.Message + "\n\n" +
                ex.GetType() +
                "\n\nStack Trace:\n" +
                ex.StackTrace;

            return MessageBox.Show(errorMessage,
                "Application Error",
                MessageBoxButtons.AbortRetryIgnore,
                MessageBoxIcon.Stop);
        }
    }

}
