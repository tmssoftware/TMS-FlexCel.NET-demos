using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Text;
using FlexCel.Render;
namespace HandlingErrors
{
    public partial class mainForm: System.Windows.Forms.Form
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.Container components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }

            //Onhook the event handler. Since this is a form, we need to onhook the event when it is disposed or it would live forever.
            FlexCelTrace.OnError -= FlexCelTrace_OnErrorHandler;

            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.label1 = new System.Windows.Forms.Label();
            this.cbStopOnErrors = new System.Windows.Forms.CheckBox();
            this.errorBox = new System.Windows.Forms.TextBox();
            this.cbIgnoreFontErrors = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.button1.Location = new System.Drawing.Point(244, 313);
            this.button1.Name = "button1";
            this.button1.TabIndex = 0;
            this.button1.Text = "GO!";
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|Excel 97/2003|*.xls|Excel 2007|*.xlsx;*.xlsm|All files|*.*";
            this.saveFileDialog1.RestoreDirectory = true;
            this.saveFileDialog1.Title = "Save file as: (FILE WILL BE SAVED AS PDF TOO)";
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(192)));
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label1.Location = new System.Drawing.Point(16, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(536, 32);
            this.label1.TabIndex = 1;
            this.label1.Text = "This demo shows how to handle non fatal errors in FlexCel by using the FlexCelTra" +
                "ce static class.";
            // 
            // cbStopOnErrors
            // 
            this.cbStopOnErrors.Location = new System.Drawing.Point(16, 64);
            this.cbStopOnErrors.Name = "cbStopOnErrors";
            this.cbStopOnErrors.Size = new System.Drawing.Size(400, 24);
            this.cbStopOnErrors.TabIndex = 2;
            this.cbStopOnErrors.Text = "Stop on non fatal errors";
            // 
            // errorBox
            // 
            this.errorBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                | System.Windows.Forms.AnchorStyles.Left)
                | System.Windows.Forms.AnchorStyles.Right)));
            this.errorBox.Font = new System.Drawing.Font("Arial Unicode MS", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.errorBox.Location = new System.Drawing.Point(16, 128);
            this.errorBox.Multiline = true;
            this.errorBox.Name = "errorBox";
            this.errorBox.ReadOnly = true;
            this.errorBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.errorBox.Size = new System.Drawing.Size(536, 160);
            this.errorBox.TabIndex = 3;
            this.errorBox.Text = "";
            this.errorBox.WordWrap = false;
            // 
            // cbIgnoreFontErrors
            // 
            this.cbIgnoreFontErrors.Location = new System.Drawing.Point(16, 88);
            this.cbIgnoreFontErrors.Name = "cbIgnoreFontErrors";
            this.cbIgnoreFontErrors.Size = new System.Drawing.Size(208, 24);
            this.cbIgnoreFontErrors.TabIndex = 4;
            this.cbIgnoreFontErrors.Text = "Ignore font errors";
            // 
            // mainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(560, 350);
            this.Controls.Add(this.cbIgnoreFontErrors);
            this.Controls.Add(this.errorBox);
            this.Controls.Add(this.cbStopOnErrors);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button1);
            this.Name = "mainForm";
            this.Text = "Handling non fatal errors.";
            this.ResumeLayout(false);

        }
        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox errorBox;
        private System.Windows.Forms.CheckBox cbStopOnErrors;
        private System.Windows.Forms.CheckBox cbIgnoreFontErrors;

    }
}

