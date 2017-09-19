using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

using System.Threading;

using FlexCel.Render;

namespace ExportingWebServices
{
    /// <summary>
    /// A dialog box to show progress. This could be done with a BackgroundWorker, it was done this way for .NET 1.1 compatibility.
    /// </summary>
    public partial class ProgressDialog: System.Windows.Forms.Form
    {
        private System.Timers.Timer timer1;

        public ProgressDialog()
        {
            InitializeComponent();
        }


        private DateTime StartTime;
        private Thread RunningThread;

        private void timer1_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            UpdateStatus();
        }

        public void ShowProgress(Thread aRunningThread)
        {
            RunningThread = aRunningThread;

            if (!RunningThread.IsAlive) { DialogResult = DialogResult.OK; return; }
            timer1.Enabled = true;
            StartTime = DateTime.Now;
            ShowDialog();
        }

        private void UpdateStatus()
        {
            TimeSpan ts = DateTime.Now - StartTime;
            string hours;
            if (ts.Hours == 0) hours = ""; else hours = ts.Hours.ToString("00") + ":";
            statusBarPanelTime.Text = hours + ts.Minutes.ToString("00") + ":" + ts.Seconds.ToString("00");

            if (!RunningThread.IsAlive) DialogResult = DialogResult.OK;
        }

        private void ProgressDialog_Closed(object sender, System.EventArgs e)
        {
            timer1.Enabled = false;
        }

    }
}
