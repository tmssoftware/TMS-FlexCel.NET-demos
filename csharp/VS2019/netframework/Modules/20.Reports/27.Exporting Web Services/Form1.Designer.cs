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
using FlexCel.Core;
using FlexCel.XlsAdapter;
using FlexCel.Report;
using FlexCel.Render;
using ExportingWebServices.gov.weather.www;
namespace ExportingWebServices
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
            this.reportStart = new FlexCel.Report.FlexCelReport();
            this.btnCancel = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.cbOffline = new System.Windows.Forms.CheckBox();
            this.btnExportPdf = new System.Windows.Forms.Button();
            this.saveFileDialogPdf = new System.Windows.Forms.SaveFileDialog();
            this.edcity = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // btnExportXls
            // 
            this.btnExportXls.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExportXls.BackColor = System.Drawing.Color.Green;
            this.btnExportXls.ForeColor = System.Drawing.Color.White;
            this.btnExportXls.Location = new System.Drawing.Point(16, 120);
            this.btnExportXls.Name = "btnExportXls";
            this.btnExportXls.Size = new System.Drawing.Size(112, 23);
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
            // reportStart
            // 
            this.reportStart.AllowOverwritingFiles = true;
            this.reportStart.DeleteEmptyRanges = false;
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.ForeColor = System.Drawing.Color.White;
            this.btnCancel.Location = new System.Drawing.Point(272, 120);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(112, 23);
            this.btnCancel.TabIndex = 3;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(32, 19);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(40, 85);
            this.label3.TabIndex = 8;
            this.label3.Text = "City:";
            // 
            // cbOffline
            // 
            this.cbOffline.Checked = true;
            this.cbOffline.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbOffline.Location = new System.Drawing.Point(32, 80);
            this.cbOffline.Name = "cbOffline";
            this.cbOffline.Size = new System.Drawing.Size(352, 24);
            this.cbOffline.TabIndex = 10;
            this.cbOffline.Text = "Use offline data. (do not actually connect to the web service)";
            // 
            // btnExportPdf
            // 
            this.btnExportPdf.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExportPdf.BackColor = System.Drawing.Color.SteelBlue;
            this.btnExportPdf.ForeColor = System.Drawing.Color.White;
            this.btnExportPdf.Location = new System.Drawing.Point(144, 120);
            this.btnExportPdf.Name = "btnExportPdf";
            this.btnExportPdf.Size = new System.Drawing.Size(112, 23);
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
            // edcity
            // 
            this.edcity.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.edcity.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.edcity.FormattingEnabled = true;
            this.edcity.Location = new System.Drawing.Point(78, 12);
            this.edcity.MaxDropDownItems = 32;
            this.edcity.Name = "edcity";
            this.edcity.Size = new System.Drawing.Size(306, 21);
            this.edcity.TabIndex = 12;
            this.edcity.KeyDown += new System.Windows.Forms.KeyEventHandler(this.edcity_KeyDown);
            // 
            // mainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(416, 157);
            this.Controls.Add(this.edcity);
            this.Controls.Add(this.btnExportPdf);
            this.Controls.Add(this.cbOffline);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnExportXls);
            this.Name = "mainForm";
            this.Text = "Exporting Web Services";
            this.ResumeLayout(false);

        }
        #endregion

        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.CheckBox cbOffline;
        private System.Windows.Forms.Button btnExportXls;
        private System.Windows.Forms.Button btnExportPdf;
        private System.Windows.Forms.SaveFileDialog saveFileDialogXls;
        private System.Windows.Forms.SaveFileDialog saveFileDialogPdf;
        private FlexCel.Report.FlexCelReport reportStart;
        private ComboBox edcity;

    }
}


