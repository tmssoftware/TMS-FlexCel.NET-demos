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
namespace HyperLinks
{
    public partial class mainForm: System.Windows.Forms.Form
    {
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.DataGrid dataGrid;
        private System.Data.DataSet dataSet1;
        private System.Data.DataTable HlDataTable;
        private System.Data.DataColumn Index;
        private System.Data.DataColumn Cell1;
        private System.Data.DataColumn Cell2;
        private System.Data.DataColumn cType;
        private System.Data.DataColumn Description;
        private System.Data.DataColumn TextMark;
        private System.Data.DataColumn TargetFrame;
        private System.Data.DataColumn cText;
        private System.Data.DataColumn cHint;
        private System.Windows.Forms.DataGridTableStyle dataGridTableStyle1;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn1;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn2;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn3;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn4;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn5;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn6;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn7;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn8;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn9;
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
            this.dataGrid = new System.Windows.Forms.DataGrid();
            this.HlDataTable = new System.Data.DataTable();
            this.Index = new System.Data.DataColumn();
            this.Cell1 = new System.Data.DataColumn();
            this.Cell2 = new System.Data.DataColumn();
            this.cType = new System.Data.DataColumn();
            this.cText = new System.Data.DataColumn();
            this.Description = new System.Data.DataColumn();
            this.TextMark = new System.Data.DataColumn();
            this.TargetFrame = new System.Data.DataColumn();
            this.cHint = new System.Data.DataColumn();
            this.dataGridTableStyle1 = new System.Windows.Forms.DataGridTableStyle();
            this.dataGridTextBoxColumn1 = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumn2 = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumn3 = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumn4 = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumn5 = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumn6 = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumn7 = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumn8 = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumn9 = new System.Windows.Forms.DataGridTextBoxColumn();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.dataSet1 = new System.Data.DataSet();
            this.mainToolbar = new System.Windows.Forms.ToolStrip();
            this.btnReadHyperlinks = new System.Windows.Forms.ToolStripButton();
            this.btnWriteHyperlinks = new System.Windows.Forms.ToolStripButton();
            this.btnExit = new System.Windows.Forms.ToolStripButton();
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.HlDataTable)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataSet1)).BeginInit();
            this.mainToolbar.SuspendLayout();
            this.SuspendLayout();
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|Excel 97/2003|*.xls|Excel 2007|*.xlsx;*.xlsm|All " +
    "files|*.*";
            this.saveFileDialog1.RestoreDirectory = true;
            this.saveFileDialog1.Title = "Save the file as...";
            // 
            // dataGrid
            // 
            this.dataGrid.CaptionText = "No file selected";
            this.dataGrid.DataMember = "";
            this.dataGrid.DataSource = this.HlDataTable;
            this.dataGrid.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGrid.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.dataGrid.Location = new System.Drawing.Point(0, 38);
            this.dataGrid.Name = "dataGrid";
            this.dataGrid.ReadOnly = true;
            this.dataGrid.Size = new System.Drawing.Size(768, 327);
            this.dataGrid.TabIndex = 3;
            this.dataGrid.TableStyles.AddRange(new System.Windows.Forms.DataGridTableStyle[] {
            this.dataGridTableStyle1});
            // 
            // HlDataTable
            // 
            this.HlDataTable.Columns.AddRange(new System.Data.DataColumn[] {
            this.Index,
            this.Cell1,
            this.Cell2,
            this.cType,
            this.cText,
            this.Description,
            this.TextMark,
            this.TargetFrame,
            this.cHint});
            this.HlDataTable.TableName = "HlDataTable";
            // 
            // Index
            // 
            this.Index.ColumnName = "Index";
            // 
            // Cell1
            // 
            this.Cell1.ColumnName = "Cell1";
            // 
            // Cell2
            // 
            this.Cell2.ColumnName = "Cell2";
            // 
            // cType
            // 
            this.cType.ColumnName = "Type";
            // 
            // cText
            // 
            this.cText.ColumnName = "Text";
            // 
            // Description
            // 
            this.Description.ColumnName = "Description";
            // 
            // TextMark
            // 
            this.TextMark.ColumnName = "TextMark";
            // 
            // TargetFrame
            // 
            this.TargetFrame.ColumnName = "TargetFrame";
            // 
            // cHint
            // 
            this.cHint.ColumnName = "Hint";
            // 
            // dataGridTableStyle1
            // 
            this.dataGridTableStyle1.DataGrid = this.dataGrid;
            this.dataGridTableStyle1.GridColumnStyles.AddRange(new System.Windows.Forms.DataGridColumnStyle[] {
            this.dataGridTextBoxColumn1,
            this.dataGridTextBoxColumn2,
            this.dataGridTextBoxColumn3,
            this.dataGridTextBoxColumn4,
            this.dataGridTextBoxColumn5,
            this.dataGridTextBoxColumn6,
            this.dataGridTextBoxColumn7,
            this.dataGridTextBoxColumn8,
            this.dataGridTextBoxColumn9});
            this.dataGridTableStyle1.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.dataGridTableStyle1.MappingName = "HlDataTable";
            this.dataGridTableStyle1.PreferredColumnWidth = 15;
            // 
            // dataGridTextBoxColumn1
            // 
            this.dataGridTextBoxColumn1.Format = "";
            this.dataGridTextBoxColumn1.FormatInfo = null;
            this.dataGridTextBoxColumn1.HeaderText = "Index";
            this.dataGridTextBoxColumn1.MappingName = "Index";
            this.dataGridTextBoxColumn1.Width = 35;
            // 
            // dataGridTextBoxColumn2
            // 
            this.dataGridTextBoxColumn2.Format = "";
            this.dataGridTextBoxColumn2.FormatInfo = null;
            this.dataGridTextBoxColumn2.HeaderText = "Cell1";
            this.dataGridTextBoxColumn2.MappingName = "Cell1";
            this.dataGridTextBoxColumn2.Width = 40;
            // 
            // dataGridTextBoxColumn3
            // 
            this.dataGridTextBoxColumn3.Format = "";
            this.dataGridTextBoxColumn3.FormatInfo = null;
            this.dataGridTextBoxColumn3.HeaderText = "Cell2";
            this.dataGridTextBoxColumn3.MappingName = "Cell2";
            this.dataGridTextBoxColumn3.Width = 40;
            // 
            // dataGridTextBoxColumn4
            // 
            this.dataGridTextBoxColumn4.Format = "";
            this.dataGridTextBoxColumn4.FormatInfo = null;
            this.dataGridTextBoxColumn4.HeaderText = "Type";
            this.dataGridTextBoxColumn4.MappingName = "Type";
            this.dataGridTextBoxColumn4.Width = 75;
            // 
            // dataGridTextBoxColumn5
            // 
            this.dataGridTextBoxColumn5.Format = "";
            this.dataGridTextBoxColumn5.FormatInfo = null;
            this.dataGridTextBoxColumn5.HeaderText = "Text";
            this.dataGridTextBoxColumn5.MappingName = "Text";
            this.dataGridTextBoxColumn5.Width = 150;
            // 
            // dataGridTextBoxColumn6
            // 
            this.dataGridTextBoxColumn6.Format = "";
            this.dataGridTextBoxColumn6.FormatInfo = null;
            this.dataGridTextBoxColumn6.HeaderText = "Description";
            this.dataGridTextBoxColumn6.MappingName = "Description";
            this.dataGridTextBoxColumn6.Width = 150;
            // 
            // dataGridTextBoxColumn7
            // 
            this.dataGridTextBoxColumn7.Format = "";
            this.dataGridTextBoxColumn7.FormatInfo = null;
            this.dataGridTextBoxColumn7.HeaderText = "TextMark";
            this.dataGridTextBoxColumn7.MappingName = "TextMark";
            this.dataGridTextBoxColumn7.Width = 75;
            // 
            // dataGridTextBoxColumn8
            // 
            this.dataGridTextBoxColumn8.Format = "";
            this.dataGridTextBoxColumn8.FormatInfo = null;
            this.dataGridTextBoxColumn8.HeaderText = "TargetFrame";
            this.dataGridTextBoxColumn8.MappingName = "TargetFrame";
            this.dataGridTextBoxColumn8.Width = 75;
            // 
            // dataGridTextBoxColumn9
            // 
            this.dataGridTextBoxColumn9.Format = "";
            this.dataGridTextBoxColumn9.FormatInfo = null;
            this.dataGridTextBoxColumn9.HeaderText = "Hint";
            this.dataGridTextBoxColumn9.MappingName = "Hint";
            this.dataGridTextBoxColumn9.Width = 75;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.DefaultExt = "xls";
            this.openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|Excel 97/2003|*.xls|Excel 2007|*.xlsx;*.xlsm|All " +
    "files|*.*";
            this.openFileDialog1.Title = "Open an Excel File";
            // 
            // dataSet1
            // 
            this.dataSet1.DataSetName = "HlDataSet";
            this.dataSet1.Locale = new System.Globalization.CultureInfo("");
            this.dataSet1.Tables.AddRange(new System.Data.DataTable[] {
            this.HlDataTable});
            // 
            // mainToolbar
            // 
            this.mainToolbar.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnReadHyperlinks,
            this.btnWriteHyperlinks,
            this.btnExit});
            this.mainToolbar.Location = new System.Drawing.Point(0, 0);
            this.mainToolbar.Name = "mainToolbar";
            this.mainToolbar.Size = new System.Drawing.Size(768, 38);
            this.mainToolbar.TabIndex = 11;
            this.mainToolbar.Text = "toolStrip1";
            // 
            // btnReadHyperlinks
            // 
            this.btnReadHyperlinks.Image = ((System.Drawing.Image)(resources.GetObject("btnReadHyperlinks.Image")));
            this.btnReadHyperlinks.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnReadHyperlinks.Name = "btnReadHyperlinks";
            this.btnReadHyperlinks.Size = new System.Drawing.Size(96, 35);
            this.btnReadHyperlinks.Text = "Read Hyperlinks";
            this.btnReadHyperlinks.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnReadHyperlinks.Click += new System.EventHandler(this.ReadHyperLinks_Click);
            // 
            // btnWriteHyperlinks
            // 
            this.btnWriteHyperlinks.Image = ((System.Drawing.Image)(resources.GetObject("btnWriteHyperlinks.Image")));
            this.btnWriteHyperlinks.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnWriteHyperlinks.Name = "btnWriteHyperlinks";
            this.btnWriteHyperlinks.Size = new System.Drawing.Size(98, 43);
            this.btnWriteHyperlinks.Text = "Write Hyperlinks";
            this.btnWriteHyperlinks.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnWriteHyperlinks.Click += new System.EventHandler(this.writeHyperLinks_Click);
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
            this.Controls.Add(this.dataGrid);
            this.Controls.Add(this.mainToolbar);
            this.Name = "mainForm";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.HlDataTable)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataSet1)).EndInit();
            this.mainToolbar.ResumeLayout(false);
            this.mainToolbar.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        private ToolStrip mainToolbar;
        private ToolStripButton btnReadHyperlinks;
        private ToolStripButton btnWriteHyperlinks;
        private ToolStripButton btnExit;
    }
}

