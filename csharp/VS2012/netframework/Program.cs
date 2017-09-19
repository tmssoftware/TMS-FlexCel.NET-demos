using System;
using System.Windows.Forms;
using System.IO;


namespace MainDemo
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.ThreadException += new System.Threading.ThreadExceptionEventHandler(Application_ThreadException);
            Application.EnableVisualStyles();
            Application.Run(new DemoForm());
        }

        private static void Application_ThreadException(object sender, System.Threading.ThreadExceptionEventArgs e)
        {
            Exception ex = e.Exception;
            while (ex.InnerException != null)
            {
                ex = ex.InnerException;
            }


            MessageBox.Show(ex.GetType().Name +" //" + ex.Message);
        }

    }
}
