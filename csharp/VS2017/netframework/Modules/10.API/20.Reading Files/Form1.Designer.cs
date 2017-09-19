using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using FlexCel.Core;
using FlexCel.XlsAdapter;
namespace ReadingFiles
{
    public partial class mainForm: System.Windows.Forms.Form
    {
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.DataGrid DisplayGrid;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox sheetCombo;
        private System.Windows.Forms.StatusBar statusBar;
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
            this.DisplayGrid = new System.Windows.Forms.DataGrid();
            this.panel1 = new System.Windows.Forms.Panel();
            this.sheetCombo = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.statusBar = new System.Windows.Forms.StatusBar();
            this.mainToolbar = new System.Windows.Forms.ToolStrip();
            this.btnOpenFile = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.btnFormatValues = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.btnValueInCellA1 = new System.Windows.Forms.ToolStripButton();
            this.btnExit = new System.Windows.Forms.ToolStripButton();
            this.btnInfo = new System.Windows.Forms.ToolStripButton();
            ((System.ComponentModel.ISupportInitialize)(this.DisplayGrid)).BeginInit();
            this.panel1.SuspendLayout();
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
            // DisplayGrid
            // 
            this.DisplayGrid.DataMember = "";
            this.DisplayGrid.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DisplayGrid.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.DisplayGrid.Location = new System.Drawing.Point(0, 67);
            this.DisplayGrid.Name = "DisplayGrid";
            this.DisplayGrid.Size = new System.Drawing.Size(880, 372);
            this.DisplayGrid.TabIndex = 5;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.ControlDark;
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.sheetCombo);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 38);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(880, 29);
            this.panel1.TabIndex = 6;
            // 
            // sheetCombo
            // 
            this.sheetCombo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.sheetCombo.Location = new System.Drawing.Point(65, 3);
            this.sheetCombo.Name = "sheetCombo";
            this.sheetCombo.Size = new System.Drawing.Size(391, 21);
            this.sheetCombo.TabIndex = 1;
            this.sheetCombo.SelectedIndexChanged += new System.EventHandler(this.sheetCombo_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.ForeColor = System.Drawing.SystemColors.HighlightText;
            this.label1.Location = new System.Drawing.Point(8, 8);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(40, 23);
            this.label1.TabIndex = 0;
            this.label1.Text = "Sheet:";
            // 
            // statusBar
            // 
            this.statusBar.Location = new System.Drawing.Point(0, 439);
            this.statusBar.Name = "statusBar";
            this.statusBar.Size = new System.Drawing.Size(880, 22);
            this.statusBar.TabIndex = 7;
            // 
            // mainToolbar
            // 
            this.mainToolbar.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnOpenFile,
            this.toolStripSeparator1,
            this.btnFormatValues,
            this.toolStripSeparator2,
            this.btnValueInCellA1,
            this.btnExit,
            this.btnInfo});
            this.mainToolbar.Location = new System.Drawing.Point(0, 0);
            this.mainToolbar.Name = "mainToolbar";
            this.mainToolbar.Size = new System.Drawing.Size(880, 38);
            this.mainToolbar.TabIndex = 11;
            this.mainToolbar.Text = "mainToolbar";
            // 
            // btnOpenFile
            // 
            this.btnOpenFile.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenFile.Image")));
            this.btnOpenFile.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnOpenFile.Name = "btnOpenFile";
            this.btnOpenFile.Size = new System.Drawing.Size(59, 35);
            this.btnOpenFile.Text = "Open file";
            this.btnOpenFile.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnOpenFile.Click += new System.EventHandler(this.btnOpenFile_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 38);
            // 
            // btnFormatValues
            // 
            this.btnFormatValues.CheckOnClick = true;
            this.btnFormatValues.Image = ((System.Drawing.Image)(resources.GetObject("btnFormatValues.Image")));
            this.btnFormatValues.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnFormatValues.Name = "btnFormatValues";
            this.btnFormatValues.Size = new System.Drawing.Size(85, 35);
            this.btnFormatValues.Text = "&Format values";
            this.btnFormatValues.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(6, 38);
            // 
            // btnValueInCellA1
            // 
            this.btnValueInCellA1.Image = ((System.Drawing.Image)(resources.GetObject("btnValueInCellA1.Image")));
            this.btnValueInCellA1.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnValueInCellA1.Name = "btnValueInCellA1";
            this.btnValueInCellA1.Size = new System.Drawing.Size(91, 35);
            this.btnValueInCellA1.Text = "&Value in cell A1";
            this.btnValueInCellA1.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnValueInCellA1.Click += new System.EventHandler(this.btnValueInCurrentCell_Click);
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
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // btnInfo
            // 
            this.btnInfo.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.btnInfo.Image = ((System.Drawing.Image)(resources.GetObject("btnInfo.Image")));
            this.btnInfo.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnInfo.Name = "btnInfo";
            this.btnInfo.Size = new System.Drawing.Size(74, 35);
            this.btnInfo.Text = "Information";
            this.btnInfo.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnInfo.Click += new System.EventHandler(this.btnInfo_Click);
            // 
            // mainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(880, 461);
            this.Controls.Add(this.DisplayGrid);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.statusBar);
            this.Controls.Add(this.mainToolbar);
            this.Name = "mainForm";
            this.Text = "Reading Excel Files";
            ((System.ComponentModel.ISupportInitialize)(this.DisplayGrid)).EndInit();
            this.panel1.ResumeLayout(false);
            this.mainToolbar.ResumeLayout(false);
            this.mainToolbar.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        private ToolStrip mainToolbar;
        private ToolStripButton btnOpenFile;
        private ToolStripSeparator toolStripSeparator1;
        private ToolStripButton btnInfo;
        private ToolStripButton btnExit;
        private ToolStripButton btnValueInCellA1;
        private ToolStripSeparator toolStripSeparator2;
        private ToolStripButton btnFormatValues;
    }
}

