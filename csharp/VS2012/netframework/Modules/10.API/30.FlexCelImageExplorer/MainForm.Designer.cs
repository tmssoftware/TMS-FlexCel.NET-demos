using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using FlexCel.Core;
using FlexCel.XlsAdapter;
namespace FlexCelImageExplorer
{
    public partial class mainForm: System.Windows.Forms.Form
    {
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Splitter splitter1;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.ListBox FilesListBox;
        private System.Windows.Forms.Label lblFolder;
        private System.Windows.Forms.DataGrid dataGrid;
        private System.Data.DataSet dataSet1;
        private System.Data.DataTable ImageDataTable;
        private System.Data.DataColumn Index;
        private System.Data.DataColumn Cell1;
        private System.Data.DataColumn Cell2;
        private System.Data.DataColumn cType;
        private System.Data.DataColumn cText;
        private System.Data.DataColumn Description;
        private System.Data.DataColumn dataColumn1;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.Splitter splitter2;
        private System.Data.DataColumn dataColumn2;
        private System.Windows.Forms.PictureBox PreviewBox;
        private System.Data.DataColumn dataColumn3;
        private System.Windows.Forms.SaveFileDialog saveImageDialog;
        private System.Data.DataColumn dataColumn4;
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.FilesListBox = new System.Windows.Forms.ListBox();
            this.panel3 = new System.Windows.Forms.Panel();
            this.lblFolder = new System.Windows.Forms.Label();
            this.splitter1 = new System.Windows.Forms.Splitter();
            this.dataGrid = new System.Windows.Forms.DataGrid();
            this.ImageDataTable = new System.Data.DataTable();
            this.dataColumn4 = new System.Data.DataColumn();
            this.Index = new System.Data.DataColumn();
            this.Cell1 = new System.Data.DataColumn();
            this.Cell2 = new System.Data.DataColumn();
            this.cType = new System.Data.DataColumn();
            this.cText = new System.Data.DataColumn();
            this.Description = new System.Data.DataColumn();
            this.dataColumn1 = new System.Data.DataColumn();
            this.dataColumn2 = new System.Data.DataColumn();
            this.dataColumn3 = new System.Data.DataColumn();
            this.dataSet1 = new System.Data.DataSet();
            this.panel4 = new System.Windows.Forms.Panel();
            this.PreviewBox = new System.Windows.Forms.PictureBox();
            this.splitter2 = new System.Windows.Forms.Splitter();
            this.saveImageDialog = new System.Windows.Forms.SaveFileDialog();
            this.mainToolbar = new System.Windows.Forms.ToolStrip();
            this.cbScanFolder = new System.Windows.Forms.ToolStripButton();
            this.btnOpenFile = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.btnShowInExcel = new System.Windows.Forms.ToolStripButton();
            this.btnSaveAsImage = new System.Windows.Forms.ToolStripButton();
            this.btnExit = new System.Windows.Forms.ToolStripButton();
            this.btnInfo = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.btnStretchPreview = new System.Windows.Forms.ToolStripButton();
            this.panel1.SuspendLayout();
            this.panel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ImageDataTable)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataSet1)).BeginInit();
            this.panel4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.PreviewBox)).BeginInit();
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
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.FilesListBox);
            this.panel1.Controls.Add(this.panel3);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel1.Location = new System.Drawing.Point(0, 38);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(176, 400);
            this.panel1.TabIndex = 4;
            // 
            // FilesListBox
            // 
            this.FilesListBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.FilesListBox.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.FilesListBox.Location = new System.Drawing.Point(0, 40);
            this.FilesListBox.Name = "FilesListBox";
            this.FilesListBox.Size = new System.Drawing.Size(172, 356);
            this.FilesListBox.TabIndex = 1;
            this.FilesListBox.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.FilesListBox_DrawItem);
            this.FilesListBox.SelectedIndexChanged += new System.EventHandler(this.FilesListBox_SelectedIndexChanged);
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.panel3.Controls.Add(this.lblFolder);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel3.Location = new System.Drawing.Point(0, 0);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(172, 40);
            this.panel3.TabIndex = 0;
            // 
            // lblFolder
            // 
            this.lblFolder.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblFolder.ForeColor = System.Drawing.Color.Black;
            this.lblFolder.Location = new System.Drawing.Point(0, 0);
            this.lblFolder.Name = "lblFolder";
            this.lblFolder.Size = new System.Drawing.Size(172, 40);
            this.lblFolder.TabIndex = 0;
            this.lblFolder.Text = "No Selected Folder.";
            // 
            // splitter1
            // 
            this.splitter1.Location = new System.Drawing.Point(176, 38);
            this.splitter1.Name = "splitter1";
            this.splitter1.Size = new System.Drawing.Size(3, 400);
            this.splitter1.TabIndex = 5;
            this.splitter1.TabStop = false;
            // 
            // dataGrid
            // 
            this.dataGrid.CaptionText = "No file selected";
            this.dataGrid.DataMember = "";
            this.dataGrid.DataSource = this.ImageDataTable;
            this.dataGrid.Dock = System.Windows.Forms.DockStyle.Top;
            this.dataGrid.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.dataGrid.Location = new System.Drawing.Point(179, 38);
            this.dataGrid.Name = "dataGrid";
            this.dataGrid.PreferredColumnWidth = 120;
            this.dataGrid.ReadOnly = true;
            this.dataGrid.Size = new System.Drawing.Size(709, 128);
            this.dataGrid.TabIndex = 7;
            // 
            // ImageDataTable
            // 
            this.ImageDataTable.Columns.AddRange(new System.Data.DataColumn[] {
            this.dataColumn4,
            this.Index,
            this.Cell1,
            this.Cell2,
            this.cType,
            this.cText,
            this.Description,
            this.dataColumn1,
            this.dataColumn2,
            this.dataColumn3});
            this.ImageDataTable.TableName = "ImageDataTable";
            // 
            // dataColumn4
            // 
            this.dataColumn4.ColumnName = "Sheet";
            // 
            // Index
            // 
            this.Index.ColumnName = "Index";
            // 
            // Cell1
            // 
            this.Cell1.Caption = "Width (Pixels)";
            this.Cell1.ColumnName = "Width (Pixels)";
            // 
            // Cell2
            // 
            this.Cell2.Caption = "Height (Pixels)";
            this.Cell2.ColumnName = "Height (Pixels)";
            // 
            // cType
            // 
            this.cType.ColumnName = "Type";
            // 
            // cText
            // 
            this.cText.Caption = "Image Format";
            this.cText.ColumnName = "Image Format";
            // 
            // Description
            // 
            this.Description.Caption = "Uncompressed size";
            this.Description.ColumnName = "Uncompressed size";
            // 
            // dataColumn1
            // 
            this.dataColumn1.Caption = "Name";
            this.dataColumn1.ColumnName = "Name";
            // 
            // dataColumn2
            // 
            this.dataColumn2.ColumnMapping = System.Data.MappingType.Hidden;
            this.dataColumn2.ColumnName = "Image";
            this.dataColumn2.DataType = typeof(byte[]);
            // 
            // dataColumn3
            // 
            this.dataColumn3.ColumnName = "Cropped";
            this.dataColumn3.DataType = typeof(bool);
            // 
            // dataSet1
            // 
            this.dataSet1.DataSetName = "ImageDataSet";
            this.dataSet1.Locale = new System.Globalization.CultureInfo("");
            this.dataSet1.Tables.AddRange(new System.Data.DataTable[] {
            this.ImageDataTable});
            // 
            // panel4
            // 
            this.panel4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel4.Controls.Add(this.PreviewBox);
            this.panel4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel4.Location = new System.Drawing.Point(179, 169);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(709, 269);
            this.panel4.TabIndex = 8;
            // 
            // PreviewBox
            // 
            this.PreviewBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.PreviewBox.Location = new System.Drawing.Point(0, 0);
            this.PreviewBox.Name = "PreviewBox";
            this.PreviewBox.Size = new System.Drawing.Size(705, 265);
            this.PreviewBox.TabIndex = 0;
            this.PreviewBox.TabStop = false;
            // 
            // splitter2
            // 
            this.splitter2.Dock = System.Windows.Forms.DockStyle.Top;
            this.splitter2.Location = new System.Drawing.Point(179, 166);
            this.splitter2.Name = "splitter2";
            this.splitter2.Size = new System.Drawing.Size(709, 3);
            this.splitter2.TabIndex = 9;
            this.splitter2.TabStop = false;
            // 
            // mainToolbar
            // 
            this.mainToolbar.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.cbScanFolder,
            this.btnOpenFile,
            this.toolStripSeparator1,
            this.btnShowInExcel,
            this.btnSaveAsImage,
            this.btnExit,
            this.btnInfo,
            this.toolStripSeparator2,
            this.btnStretchPreview});
            this.mainToolbar.Location = new System.Drawing.Point(0, 0);
            this.mainToolbar.Name = "mainToolbar";
            this.mainToolbar.Size = new System.Drawing.Size(888, 38);
            this.mainToolbar.TabIndex = 12;
            this.mainToolbar.Text = "toolStrip1";
            // 
            // cbScanFolder
            // 
            this.cbScanFolder.CheckOnClick = true;
            this.cbScanFolder.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.cbScanFolder.Image = ((System.Drawing.Image)(resources.GetObject("cbScanFolder.Image")));
            this.cbScanFolder.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.cbScanFolder.Name = "cbScanFolder";
            this.cbScanFolder.Size = new System.Drawing.Size(122, 35);
            this.cbScanFolder.Text = "Scan all files in folder";
            this.cbScanFolder.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
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
            // btnShowInExcel
            // 
            this.btnShowInExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnShowInExcel.Image")));
            this.btnShowInExcel.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnShowInExcel.Name = "btnShowInExcel";
            this.btnShowInExcel.Size = new System.Drawing.Size(82, 35);
            this.btnShowInExcel.Text = "Show in Excel";
            this.btnShowInExcel.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnShowInExcel.Click += new System.EventHandler(this.btnOpen_Click);
            // 
            // btnSaveAsImage
            // 
            this.btnSaveAsImage.Image = ((System.Drawing.Image)(resources.GetObject("btnSaveAsImage.Image")));
            this.btnSaveAsImage.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnSaveAsImage.Name = "btnSaveAsImage";
            this.btnSaveAsImage.Size = new System.Drawing.Size(85, 35);
            this.btnSaveAsImage.Text = "Save as image";
            this.btnSaveAsImage.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnSaveAsImage.Click += new System.EventHandler(this.btnSaveImage_Click);
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
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(6, 38);
            // 
            // btnStretchPreview
            // 
            this.btnStretchPreview.CheckOnClick = true;
            this.btnStretchPreview.Image = ((System.Drawing.Image)(resources.GetObject("btnStretchPreview.Image")));
            this.btnStretchPreview.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnStretchPreview.Name = "btnStretchPreview";
            this.btnStretchPreview.Size = new System.Drawing.Size(92, 35);
            this.btnStretchPreview.Text = "Stretch preview";
            this.btnStretchPreview.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnStretchPreview.Click += new System.EventHandler(this.btnStretchPreview_Click);
            // 
            // mainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(888, 438);
            this.Controls.Add(this.panel4);
            this.Controls.Add(this.splitter2);
            this.Controls.Add(this.dataGrid);
            this.Controls.Add(this.splitter1);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.mainToolbar);
            this.Name = "mainForm";
            this.Text = "FlexCel Image Explorer";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.panel1.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ImageDataTable)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataSet1)).EndInit();
            this.panel4.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.PreviewBox)).EndInit();
            this.mainToolbar.ResumeLayout(false);
            this.mainToolbar.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        private ToolStrip mainToolbar;
        private ToolStripButton btnOpenFile;
        private ToolStripSeparator toolStripSeparator1;
        private ToolStripButton btnShowInExcel;
        private ToolStripButton btnSaveAsImage;
        private ToolStripButton btnExit;
        private ToolStripButton btnInfo;
        private ToolStripSeparator toolStripSeparator2;
        private ToolStripButton btnStretchPreview;
        private ToolStripButton cbScanFolder;
    }
}

