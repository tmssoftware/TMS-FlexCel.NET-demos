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
namespace ValidateRecalc
{
    public partial class mainForm: System.Windows.Forms.Form
    {
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.RichTextBox report;
        private System.Windows.Forms.OpenFileDialog linkedFileDialog;
        private System.ComponentModel.IContainer components = null;

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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(mainForm));
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.report = new System.Windows.Forms.RichTextBox();
            this.XlsReport = new FlexCel.Report.FlexCelReport();
            this.linkedFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.mainToolbar = new System.Windows.Forms.ToolStrip();
            this.validateRecalc = new System.Windows.Forms.ToolStripButton();
            this.compareWithExcel = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.btnInfo = new System.Windows.Forms.ToolStripButton();
            this.btnExit = new System.Windows.Forms.ToolStripButton();
            this.mainToolbar.SuspendLayout();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.DefaultExt = "xls";
            this.openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|Excel 97/2003|*.xls|Excel 2007|*.xlsx;*.xlsm|All " +
    "files|*.*";
            this.openFileDialog1.Title = "Open an Excel File";
            // 
            // report
            // 
            this.report.Dock = System.Windows.Forms.DockStyle.Fill;
            this.report.Font = new System.Drawing.Font("Courier New", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.report.Location = new System.Drawing.Point(0, 38);
            this.report.Name = "report";
            this.report.ReadOnly = true;
            this.report.Size = new System.Drawing.Size(768, 327);
            this.report.TabIndex = 3;
            this.report.Text = "";
            // 
            // linkedFileDialog
            // 
            this.linkedFileDialog.DefaultExt = "xls";
            this.linkedFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|Excel 97/2003|*.xls|Excel 2007|*.xlsx;*.xlsm|All " +
    "files|*.*";
            this.linkedFileDialog.Title = "Please supply the location for the following linked file.";
            // 
            // mainToolbar
            // 
            this.mainToolbar.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.validateRecalc,
            this.compareWithExcel,
            this.toolStripSeparator1,
            this.btnInfo,
            this.btnExit});
            this.mainToolbar.Location = new System.Drawing.Point(0, 0);
            this.mainToolbar.Name = "mainToolbar";
            this.mainToolbar.Size = new System.Drawing.Size(768, 38);
            this.mainToolbar.TabIndex = 11;
            this.mainToolbar.Text = "toolStrip1";
            // 
            // validateRecalc
            // 
            this.validateRecalc.Image = ((System.Drawing.Image)(resources.GetObject("validateRecalc.Image")));
            this.validateRecalc.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.validateRecalc.Name = "validateRecalc";
            this.validateRecalc.Size = new System.Drawing.Size(90, 35);
            this.validateRecalc.Text = "&Validate Recalc";
            this.validateRecalc.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.validateRecalc.Click += new System.EventHandler(this.validateRecalc_Click);
            // 
            // compareWithExcel
            // 
            this.compareWithExcel.Image = ((System.Drawing.Image)(resources.GetObject("compareWithExcel.Image")));
            this.compareWithExcel.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.compareWithExcel.Name = "compareWithExcel";
            this.compareWithExcel.Size = new System.Drawing.Size(115, 43);
            this.compareWithExcel.Text = "Compare with Excel";
            this.compareWithExcel.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.compareWithExcel.Click += new System.EventHandler(this.compareWithExcel_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 46);
            // 
            // btnInfo
            // 
            this.btnInfo.Image = ((System.Drawing.Image)(resources.GetObject("btnInfo.Image")));
            this.btnInfo.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnInfo.Name = "btnInfo";
            this.btnInfo.Size = new System.Drawing.Size(74, 43);
            this.btnInfo.Text = "Information";
            this.btnInfo.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnInfo.Click += new System.EventHandler(this.btnInfo_Click);
            // 
            // btnExit
            // 
            this.btnExit.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.btnExit.Image = ((System.Drawing.Image)(resources.GetObject("btnExit.Image")));
            this.btnExit.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(59, 35);
            this.btnExit.Text = "     E&xit     ";
            this.btnExit.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnExit.Click += new System.EventHandler(this.button2_Click);
            // 
            // mainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(768, 365);
            this.Controls.Add(this.report);
            this.Controls.Add(this.mainToolbar);
            this.Name = "mainForm";
            this.Text = "Validate FlexCel recalculation";
            this.mainToolbar.ResumeLayout(false);
            this.mainToolbar.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        private ToolStrip mainToolbar;
        private ToolStripButton validateRecalc;
        private ToolStripButton compareWithExcel;
        private ToolStripSeparator toolStripSeparator1;
        private ToolStripButton btnInfo;
        private ToolStripButton btnExit;
    }
}

