using System;
using System.Drawing;
using System.ComponentModel;
using System.Windows.Forms;

using System.Threading;

using FlexCel.Render;

namespace CustomPreview
{
    /// <summary>
    /// Shows progress as we are exporting to pdf.
    /// </summary>
    public partial class PdfProgressDialog: System.Windows.Forms.Form
    {
        private System.Timers.Timer timer1;

        public PdfProgressDialog()
        {
            InitializeComponent();
        }


        private DateTime StartTime;
        private Thread RunningThread;
        private FlexCelPdfExport PdfExport;

        private void timer1_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            UpdateStatus();
        }

        public void ShowProgress(Thread aRunningThread, FlexCelPdfExport aPdfExport)
        {
            RunningThread = aRunningThread;

            if (!RunningThread.IsAlive) { DialogResult = DialogResult.OK; return; }
            timer1.Enabled = true;
            StartTime = DateTime.Now;
            PdfExport = aPdfExport;
            ShowDialog();
        }

        private void UpdateStatus()
        {
            TimeSpan ts = DateTime.Now - StartTime;
            string hours;
            if (ts.Hours == 0) hours = ""; else hours = ts.Hours.ToString("00") + ":";
            statusBarPanelTime.Text = hours + ts.Minutes.ToString("00") + ":" + ts.Seconds.ToString("00");

            if (!RunningThread.IsAlive) DialogResult = DialogResult.OK;

            if (PdfExport.Progress.TotalPage > 0) labelPages.Text = String.Format("Generating Page {0} of {1}", PdfExport.Progress.Page, PdfExport.Progress.TotalPage);
        }

        private void PdfProgressDialog_Closed(object sender, System.EventArgs e)
        {
            timer1.Enabled = false;
        }


    }
}
