using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Reflection;
using FlexCel.Core;

namespace MainDemo
{
    /// <summary>
    /// About...
    /// </summary>
    public partial class AboutForm: System.Windows.Forms.Form
    {

        public AboutForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            Close();
        }

        private void linkLabel1_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(linkLabel1.Text);
        }

        private void linkLabel2_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(linkLabel2.Text);
        }

        private void AboutForm_Load(object sender, System.EventArgs e)
        {
            Assembly asm = Assembly.GetAssembly(typeof(ExcelFile));
            lblVersion.Text = "Using FlexCel Version: " + asm.GetName().Version.ToString();

        }
    }
}
