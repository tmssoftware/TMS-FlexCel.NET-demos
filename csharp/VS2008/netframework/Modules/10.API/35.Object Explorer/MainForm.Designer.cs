using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using FlexCel.Core;
using FlexCel.XlsAdapter;
namespace ObjectExplorer
{
    public partial class mainForm: System.Windows.Forms.Form
    {
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Splitter splitter1;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.DataGrid dataGrid;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.Splitter splitter2;
        private System.Windows.Forms.PictureBox PreviewBox;
        private System.Windows.Forms.SaveFileDialog saveImageDialog;
        private System.Windows.Forms.Label lblObjects;
        private System.Windows.Forms.TreeView ObjTree;
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.Label lblObjName;
        private System.Windows.Forms.Label lblObjText;
        private System.Windows.Forms.Panel panel6;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cbSheet;
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
            this.ObjTree = new System.Windows.Forms.TreeView();
            this.panel6 = new System.Windows.Forms.Panel();
            this.cbSheet = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.panel3 = new System.Windows.Forms.Panel();
            this.lblObjects = new System.Windows.Forms.Label();
            this.splitter1 = new System.Windows.Forms.Splitter();
            this.dataGrid = new System.Windows.Forms.DataGrid();
            this.panel4 = new System.Windows.Forms.Panel();
            this.PreviewBox = new System.Windows.Forms.PictureBox();
            this.splitter2 = new System.Windows.Forms.Splitter();
            this.saveImageDialog = new System.Windows.Forms.SaveFileDialog();
            this.panel5 = new System.Windows.Forms.Panel();
            this.lblObjText = new System.Windows.Forms.Label();
            this.lblObjName = new System.Windows.Forms.Label();
            this.mainToolbar = new System.Windows.Forms.ToolStrip();
            this.btnOpenFile = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.btnShowInExcel = new System.Windows.Forms.ToolStripButton();
            this.btnSaveAsImage = new System.Windows.Forms.ToolStripButton();
            this.btnExit = new System.Windows.Forms.ToolStripButton();
            this.btnInfo = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.btnStretchPreview = new System.Windows.Forms.ToolStripButton();
            this.panel1.SuspendLayout();
            this.panel6.SuspendLayout();
            this.panel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid)).BeginInit();
            this.panel4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.PreviewBox)).BeginInit();
            this.panel5.SuspendLayout();
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
            this.panel1.Controls.Add(this.ObjTree);
            this.panel1.Controls.Add(this.panel6);
            this.panel1.Controls.Add(this.panel3);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel1.Location = new System.Drawing.Point(0, 38);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(176, 472);
            this.panel1.TabIndex = 4;
            // 
            // ObjTree
            // 
            this.ObjTree.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ObjTree.HideSelection = false;
            this.ObjTree.Location = new System.Drawing.Point(0, 112);
            this.ObjTree.Name = "ObjTree";
            this.ObjTree.Size = new System.Drawing.Size(172, 356);
            this.ObjTree.TabIndex = 1;
            this.ObjTree.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.ObjTree_AfterSelect);
            // 
            // panel6
            // 
            this.panel6.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.panel6.Controls.Add(this.cbSheet);
            this.panel6.Controls.Add(this.label1);
            this.panel6.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel6.Location = new System.Drawing.Point(0, 64);
            this.panel6.Name = "panel6";
            this.panel6.Size = new System.Drawing.Size(172, 48);
            this.panel6.TabIndex = 2;
            // 
            // cbSheet
            // 
            this.cbSheet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbSheet.Location = new System.Drawing.Point(6, 16);
            this.cbSheet.Name = "cbSheet";
            this.cbSheet.Size = new System.Drawing.Size(160, 21);
            this.cbSheet.TabIndex = 33;
            this.cbSheet.SelectedIndexChanged += new System.EventHandler(this.cbSheet_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.LightGoldenrodYellow;
            this.label1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label1.ForeColor = System.Drawing.Color.Black;
            this.label1.Location = new System.Drawing.Point(0, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(172, 48);
            this.label1.TabIndex = 1;
            this.label1.Text = "Sheet:";
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.panel3.Controls.Add(this.lblObjects);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel3.Location = new System.Drawing.Point(0, 0);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(172, 64);
            this.panel3.TabIndex = 0;
            // 
            // lblObjects
            // 
            this.lblObjects.BackColor = System.Drawing.Color.LightGoldenrodYellow;
            this.lblObjects.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblObjects.ForeColor = System.Drawing.Color.Black;
            this.lblObjects.Location = new System.Drawing.Point(0, 0);
            this.lblObjects.Name = "lblObjects";
            this.lblObjects.Size = new System.Drawing.Size(172, 64);
            this.lblObjects.TabIndex = 1;
            this.lblObjects.Text = "File:";
            // 
            // splitter1
            // 
            this.splitter1.Location = new System.Drawing.Point(176, 91);
            this.splitter1.Name = "splitter1";
            this.splitter1.Size = new System.Drawing.Size(3, 419);
            this.splitter1.TabIndex = 5;
            this.splitter1.TabStop = false;
            // 
            // dataGrid
            // 
            this.dataGrid.CaptionText = "Object Properties";
            this.dataGrid.DataMember = "";
            this.dataGrid.Dock = System.Windows.Forms.DockStyle.Top;
            this.dataGrid.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.dataGrid.Location = new System.Drawing.Point(179, 91);
            this.dataGrid.Name = "dataGrid";
            this.dataGrid.PreferredColumnWidth = 120;
            this.dataGrid.ReadOnly = true;
            this.dataGrid.Size = new System.Drawing.Size(557, 213);
            this.dataGrid.TabIndex = 7;
            // 
            // panel4
            // 
            this.panel4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel4.Controls.Add(this.PreviewBox);
            this.panel4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel4.Location = new System.Drawing.Point(179, 307);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(557, 203);
            this.panel4.TabIndex = 8;
            // 
            // PreviewBox
            // 
            this.PreviewBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.PreviewBox.Location = new System.Drawing.Point(0, 0);
            this.PreviewBox.Name = "PreviewBox";
            this.PreviewBox.Size = new System.Drawing.Size(553, 199);
            this.PreviewBox.TabIndex = 0;
            this.PreviewBox.TabStop = false;
            // 
            // splitter2
            // 
            this.splitter2.Dock = System.Windows.Forms.DockStyle.Top;
            this.splitter2.Location = new System.Drawing.Point(179, 304);
            this.splitter2.Name = "splitter2";
            this.splitter2.Size = new System.Drawing.Size(557, 3);
            this.splitter2.TabIndex = 9;
            this.splitter2.TabStop = false;
            // 
            // saveImageDialog
            // 
            this.saveImageDialog.DefaultExt = "png";
            this.saveImageDialog.Filter = "PNG Files|*.png";
            // 
            // panel5
            // 
            this.panel5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel5.Controls.Add(this.lblObjText);
            this.panel5.Controls.Add(this.lblObjName);
            this.panel5.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel5.Location = new System.Drawing.Point(176, 38);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(560, 53);
            this.panel5.TabIndex = 10;
            // 
            // lblObjText
            // 
            this.lblObjText.AutoSize = true;
            this.lblObjText.Location = new System.Drawing.Point(16, 32);
            this.lblObjText.Name = "lblObjText";
            this.lblObjText.Size = new System.Drawing.Size(31, 13);
            this.lblObjText.TabIndex = 1;
            this.lblObjText.Text = "Text:";
            // 
            // lblObjName
            // 
            this.lblObjName.AutoSize = true;
            this.lblObjName.Location = new System.Drawing.Point(16, 8);
            this.lblObjName.Name = "lblObjName";
            this.lblObjName.Size = new System.Drawing.Size(38, 13);
            this.lblObjName.TabIndex = 0;
            this.lblObjName.Text = "Name:";
            // 
            // mainToolbar
            // 
            this.mainToolbar.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
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
            this.mainToolbar.Size = new System.Drawing.Size(736, 38);
            this.mainToolbar.TabIndex = 11;
            this.mainToolbar.Text = "toolStrip1";
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
            this.ClientSize = new System.Drawing.Size(736, 510);
            this.Controls.Add(this.panel4);
            this.Controls.Add(this.splitter2);
            this.Controls.Add(this.dataGrid);
            this.Controls.Add(this.splitter1);
            this.Controls.Add(this.panel5);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.mainToolbar);
            this.Name = "mainForm";
            this.Text = "FlexCel Object Explorer";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.panel1.ResumeLayout(false);
            this.panel6.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid)).EndInit();
            this.panel4.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.PreviewBox)).EndInit();
            this.panel5.ResumeLayout(false);
            this.panel5.PerformLayout();
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
        private ToolStripButton btnExit;
        private ToolStripButton btnSaveAsImage;
        private ToolStripButton btnInfo;
        private ToolStripSeparator toolStripSeparator2;
        private ToolStripButton btnStretchPreview;


    }
}

