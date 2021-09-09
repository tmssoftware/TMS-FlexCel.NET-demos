using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Data.OleDb;
using System.Threading;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using FlexCel.Report;
using System.Xml;
namespace MetaTemplates
{
    public partial class mainForm: System.Windows.Forms.Form
    {
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button btnExportExcel;
        private System.Windows.Forms.ComboBox cbFeeds;
        private System.Windows.Forms.CheckBox cbOffline;
        private System.Windows.Forms.CheckBox cbShowFeedCount;
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
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.panel2 = new System.Windows.Forms.Panel();
            this.button2 = new System.Windows.Forms.Button();
            this.btnExportExcel = new System.Windows.Forms.Button();
            this.cbFeeds = new System.Windows.Forms.ComboBox();
            this.cbOffline = new System.Windows.Forms.CheckBox();
            this.cbShowFeedCount = new System.Windows.Forms.CheckBox();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|Excel 97/2003|*.xls|Excel 2007|*.xlsx;*.xlsm|All " +
    "files|*.*";
            this.saveFileDialog1.RestoreDirectory = true;
            // 
            // panel2
            // 
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel2.Controls.Add(this.button2);
            this.panel2.Controls.Add(this.btnExportExcel);
            this.panel2.Controls.Add(this.cbFeeds);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(528, 40);
            this.panel2.TabIndex = 3;
            // 
            // button2
            // 
            this.button2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button2.BackColor = System.Drawing.SystemColors.Control;
            this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
            this.button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button2.Location = new System.Drawing.Point(464, 2);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(56, 26);
            this.button2.TabIndex = 2;
            this.button2.Text = "Exit";
            this.button2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // btnExportExcel
            // 
            this.btnExportExcel.BackColor = System.Drawing.SystemColors.Control;
            this.btnExportExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExportExcel.Image")));
            this.btnExportExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnExportExcel.Location = new System.Drawing.Point(16, 2);
            this.btnExportExcel.Name = "btnExportExcel";
            this.btnExportExcel.Size = new System.Drawing.Size(120, 30);
            this.btnExportExcel.TabIndex = 1;
            this.btnExportExcel.Text = "Export to Excel";
            this.btnExportExcel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnExportExcel.UseVisualStyleBackColor = false;
            this.btnExportExcel.Click += new System.EventHandler(this.btnExportExcel_Click);
            // 
            // cbFeeds
            // 
            this.cbFeeds.DisplayMember = "1";
            this.cbFeeds.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbFeeds.Location = new System.Drawing.Point(200, 8);
            this.cbFeeds.Name = "cbFeeds";
            this.cbFeeds.Size = new System.Drawing.Size(216, 21);
            this.cbFeeds.TabIndex = 4;
            this.cbFeeds.ValueMember = "1";
            // 
            // cbOffline
            // 
            this.cbOffline.Checked = true;
            this.cbOffline.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbOffline.Location = new System.Drawing.Point(40, 56);
            this.cbOffline.Name = "cbOffline";
            this.cbOffline.Size = new System.Drawing.Size(368, 24);
            this.cbOffline.TabIndex = 5;
            this.cbOffline.Text = "Use offline data (do not connect to internet)";
            // 
            // cbShowFeedCount
            // 
            this.cbShowFeedCount.Location = new System.Drawing.Point(40, 80);
            this.cbShowFeedCount.Name = "cbShowFeedCount";
            this.cbShowFeedCount.Size = new System.Drawing.Size(360, 24);
            this.cbShowFeedCount.TabIndex = 6;
            this.cbShowFeedCount.Text = "Show feed number column in the generated report.";
            // 
            // mainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(528, 126);
            this.Controls.Add(this.cbShowFeedCount);
            this.Controls.Add(this.cbOffline);
            this.Controls.Add(this.panel2);
            this.Name = "mainForm";
            this.Text = "Meta Templates";
            this.Load += new System.EventHandler(this.mainForm_Load);
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);

        }
        #endregion
    }
}

