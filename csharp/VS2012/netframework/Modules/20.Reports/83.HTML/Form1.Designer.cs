using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Xml;
using System.Net;
using System.Threading;
using System.Globalization;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using FlexCel.Report;
using FlexCel.Render;
namespace HTML
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
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnExportXls = new System.Windows.Forms.Button();
            this.saveFileDialogXls = new System.Windows.Forms.SaveFileDialog();
            this.btnCancel = new System.Windows.Forms.Button();
            this.cbOffline = new System.Windows.Forms.CheckBox();
            this.btnExportPdf = new System.Windows.Forms.Button();
            this.saveFileDialogPdf = new System.Windows.Forms.SaveFileDialog();
            this.edCity = new System.Windows.Forms.TextBox();
            this.labelCity = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnExportXls
            // 
            this.btnExportXls.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExportXls.BackColor = System.Drawing.Color.Green;
            this.btnExportXls.ForeColor = System.Drawing.Color.White;
            this.btnExportXls.Location = new System.Drawing.Point(16, 258);
            this.btnExportXls.Name = "btnExportXls";
            this.btnExportXls.Size = new System.Drawing.Size(112, 24);
            this.btnExportXls.TabIndex = 0;
            this.btnExportXls.Text = "Export to Excel";
            this.btnExportXls.UseVisualStyleBackColor = false;
            this.btnExportXls.Click += new System.EventHandler(this.btnExportXls_Click);
            // 
            // saveFileDialogXls
            // 
            this.saveFileDialogXls.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|Excel 97/2003|*.xls|Excel 2007|*.xlsx;*.xlsm|All " +
    "files|*.*";
            this.saveFileDialogXls.RestoreDirectory = true;
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.ForeColor = System.Drawing.Color.White;
            this.btnCancel.Location = new System.Drawing.Point(272, 258);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(112, 24);
            this.btnCancel.TabIndex = 3;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // cbOffline
            // 
            this.cbOffline.Checked = true;
            this.cbOffline.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbOffline.Location = new System.Drawing.Point(13, 200);
            this.cbOffline.Name = "cbOffline";
            this.cbOffline.Size = new System.Drawing.Size(352, 24);
            this.cbOffline.TabIndex = 10;
            this.cbOffline.Text = "Use offline data. (do not actually connect to the web service)";
            // 
            // btnExportPdf
            // 
            this.btnExportPdf.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExportPdf.BackColor = System.Drawing.Color.SteelBlue;
            this.btnExportPdf.ForeColor = System.Drawing.Color.White;
            this.btnExportPdf.Location = new System.Drawing.Point(144, 258);
            this.btnExportPdf.Name = "btnExportPdf";
            this.btnExportPdf.Size = new System.Drawing.Size(112, 24);
            this.btnExportPdf.TabIndex = 11;
            this.btnExportPdf.Text = "Export to Pdf";
            this.btnExportPdf.UseVisualStyleBackColor = false;
            this.btnExportPdf.Click += new System.EventHandler(this.btnExportPdf_Click);
            // 
            // saveFileDialogPdf
            // 
            this.saveFileDialogPdf.Filter = "Pdf Files|*.pdf";
            this.saveFileDialogPdf.RestoreDirectory = true;
            // 
            // edCity
            // 
            this.edCity.Location = new System.Drawing.Point(157, 160);
            this.edCity.Name = "edCity";
            this.edCity.Size = new System.Drawing.Size(208, 20);
            this.edCity.TabIndex = 12;
            this.edCity.Text = "london";
            // 
            // labelCity
            // 
            this.labelCity.Location = new System.Drawing.Point(5, 152);
            this.labelCity.Name = "labelCity";
            this.labelCity.Size = new System.Drawing.Size(144, 40);
            this.labelCity.TabIndex = 13;
            this.labelCity.Text = "City Name: (try things like tokio, sydney, new york, madrid, rio de janeiro)";
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Times New Roman", 10.25F, System.Drawing.FontStyle.Italic);
            this.label1.Location = new System.Drawing.Point(13, 88);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(352, 32);
            this.label1.TabIndex = 14;
            this.label1.Text = "This application uses Yahoo Travel APIs to load demo trips and export them to Exc" +
    "el. For more information, visit:";
            // 
            // linkLabel1
            // 
            this.linkLabel1.Location = new System.Drawing.Point(13, 128);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(320, 16);
            this.linkLabel1.TabIndex = 15;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "http://developer.yahoo.com/travel/";
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.LightGoldenrodYellow;
            this.panel1.Controls.Add(this.label2);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(416, 69);
            this.panel1.TabIndex = 16;
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(12, 9);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(392, 49);
            this.label2.TabIndex = 0;
            this.label2.Text = "IMPORTANT: Yahoo has discontinued this service, so this demo will only work with " +
    "offline data. As the online functionality isn\'t essential to this demo, we have " +
    "decided to keep it.";
            // 
            // mainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(416, 292);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.linkLabel1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.labelCity);
            this.Controls.Add(this.edCity);
            this.Controls.Add(this.btnExportPdf);
            this.Controls.Add(this.cbOffline);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnExportXls);
            this.Name = "mainForm";
            this.Text = "Using HTML formatted text with FlexCel";
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.CheckBox cbOffline;
        private System.Windows.Forms.Button btnExportXls;
        private System.Windows.Forms.Button btnExportPdf;
        private System.Windows.Forms.SaveFileDialog saveFileDialogXls;
        private System.Windows.Forms.SaveFileDialog saveFileDialogPdf;
        private System.Windows.Forms.TextBox edCity;
        private System.Windows.Forms.Label labelCity;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private Panel panel1;
        private Label label2;
    }
}

