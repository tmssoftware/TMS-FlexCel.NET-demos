using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using FlexCel.Core;
using FlexCel.XlsAdapter;
namespace VirtualMode
{
    public partial class mainForm: System.Windows.Forms.Form
    {
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Panel panel2;
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(mainForm));
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.panel2 = new System.Windows.Forms.Panel();
            this.cbFormatValues = new System.Windows.Forms.CheckBox();
            this.cbIgnoreFormulaText = new System.Windows.Forms.CheckBox();
            this.cbFirst50Rows = new System.Windows.Forms.CheckBox();
            this.statusBar = new System.Windows.Forms.StatusBar();
            this.DisplayGrid = new System.Windows.Forms.DataGridView();
            this.GridCaptionPanel = new System.Windows.Forms.Panel();
            this.GridCaption = new System.Windows.Forms.Label();
            this.mainToolbar = new System.Windows.Forms.ToolStrip();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.btnOpenFile = new System.Windows.Forms.ToolStripButton();
            this.btnExit = new System.Windows.Forms.ToolStripButton();
            this.btnInfo = new System.Windows.Forms.ToolStripButton();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DisplayGrid)).BeginInit();
            this.GridCaptionPanel.SuspendLayout();
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
            // panel2
            // 
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel2.Controls.Add(this.cbFormatValues);
            this.panel2.Controls.Add(this.cbIgnoreFormulaText);
            this.panel2.Controls.Add(this.cbFirst50Rows);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel2.Location = new System.Drawing.Point(0, 38);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(880, 25);
            this.panel2.TabIndex = 4;
            // 
            // cbFormatValues
            // 
            this.cbFormatValues.AutoSize = true;
            this.cbFormatValues.Location = new System.Drawing.Point(269, 5);
            this.cbFormatValues.Name = "cbFormatValues";
            this.cbFormatValues.Size = new System.Drawing.Size(131, 17);
            this.cbFormatValues.TabIndex = 2;
            this.cbFormatValues.Text = "Format values (slower)";
            this.cbFormatValues.UseVisualStyleBackColor = true;
            // 
            // cbIgnoreFormulaText
            // 
            this.cbIgnoreFormulaText.AutoSize = true;
            this.cbIgnoreFormulaText.Checked = true;
            this.cbIgnoreFormulaText.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbIgnoreFormulaText.Location = new System.Drawing.Point(150, 5);
            this.cbIgnoreFormulaText.Name = "cbIgnoreFormulaText";
            this.cbIgnoreFormulaText.Size = new System.Drawing.Size(113, 17);
            this.cbIgnoreFormulaText.TabIndex = 1;
            this.cbIgnoreFormulaText.Text = "Ignore formula text";
            this.cbIgnoreFormulaText.UseVisualStyleBackColor = true;
            // 
            // cbFirst50Rows
            // 
            this.cbFirst50Rows.AutoSize = true;
            this.cbFirst50Rows.Checked = true;
            this.cbFirst50Rows.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbFirst50Rows.Location = new System.Drawing.Point(11, 5);
            this.cbFirst50Rows.Name = "cbFirst50Rows";
            this.cbFirst50Rows.Size = new System.Drawing.Size(133, 17);
            this.cbFirst50Rows.TabIndex = 0;
            this.cbFirst50Rows.Text = "Read only first 50 rows";
            this.cbFirst50Rows.UseVisualStyleBackColor = true;
            // 
            // statusBar
            // 
            this.statusBar.Location = new System.Drawing.Point(0, 439);
            this.statusBar.Name = "statusBar";
            this.statusBar.Size = new System.Drawing.Size(880, 22);
            this.statusBar.TabIndex = 7;
            // 
            // DisplayGrid
            // 
            this.DisplayGrid.AllowUserToAddRows = false;
            this.DisplayGrid.AllowUserToDeleteRows = false;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.DisplayGrid.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.DisplayGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DisplayGrid.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DisplayGrid.Location = new System.Drawing.Point(0, 86);
            this.DisplayGrid.Name = "DisplayGrid";
            this.DisplayGrid.ReadOnly = true;
            this.DisplayGrid.Size = new System.Drawing.Size(880, 353);
            this.DisplayGrid.TabIndex = 8;
            this.DisplayGrid.VirtualMode = true;
            this.DisplayGrid.CellValueNeeded += new System.Windows.Forms.DataGridViewCellValueEventHandler(this.DisplayGrid_CellValueNeeded);
            this.DisplayGrid.RowPostPaint += new System.Windows.Forms.DataGridViewRowPostPaintEventHandler(this.DisplayGrid_RowPostPaint);
            // 
            // GridCaptionPanel
            // 
            this.GridCaptionPanel.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.GridCaptionPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.GridCaptionPanel.Controls.Add(this.GridCaption);
            this.GridCaptionPanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.GridCaptionPanel.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.GridCaptionPanel.Location = new System.Drawing.Point(0, 63);
            this.GridCaptionPanel.Name = "GridCaptionPanel";
            this.GridCaptionPanel.Size = new System.Drawing.Size(880, 23);
            this.GridCaptionPanel.TabIndex = 9;
            // 
            // GridCaption
            // 
            this.GridCaption.AutoSize = true;
            this.GridCaption.Location = new System.Drawing.Point(13, 6);
            this.GridCaption.Name = "GridCaption";
            this.GridCaption.Size = new System.Drawing.Size(0, 13);
            this.GridCaption.TabIndex = 0;
            // 
            // mainToolbar
            // 
            this.mainToolbar.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnOpenFile,
            this.toolStripSeparator1,
            this.btnExit,
            this.btnInfo});
            this.mainToolbar.Location = new System.Drawing.Point(0, 0);
            this.mainToolbar.Name = "mainToolbar";
            this.mainToolbar.Size = new System.Drawing.Size(880, 38);
            this.mainToolbar.TabIndex = 10;
            this.mainToolbar.Text = "toolStrip1";
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 38);
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
            this.Controls.Add(this.GridCaptionPanel);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.statusBar);
            this.Controls.Add(this.mainToolbar);
            this.Name = "mainForm";
            this.Text = "Virtual Mode Example - Cells are not stored in memory";
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DisplayGrid)).EndInit();
            this.GridCaptionPanel.ResumeLayout(false);
            this.GridCaptionPanel.PerformLayout();
            this.mainToolbar.ResumeLayout(false);
            this.mainToolbar.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        private DataGridView DisplayGrid;
        private Panel GridCaptionPanel;
        private Label GridCaption;
        private ToolStrip mainToolbar;
        private ToolStripButton btnOpenFile;
        private ToolStripSeparator toolStripSeparator1;
        private ToolStripButton btnInfo;
        private ToolStripButton btnExit;
        private CheckBox cbFirst50Rows;
        private CheckBox cbIgnoreFormulaText;
        private CheckBox cbFormatValues;
    }
}

