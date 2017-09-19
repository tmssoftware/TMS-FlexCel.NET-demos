using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using FlexCel.Render;
using System.IO;
using System.Diagnostics;
using System.Text;
namespace ExportSVG
{
    public partial class mainForm: System.Windows.Forms.Form
    {
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.TextBox edTop;
        private System.Windows.Forms.TextBox edLeft;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.TextBox edRight;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.TextBox edBottom;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.SaveFileDialog exportDialog;
        private System.Windows.Forms.Panel panel8;
        private System.Windows.Forms.CheckBox chFormulaText;
        private System.Windows.Forms.CheckBox chGridLines;
        private System.Windows.Forms.Label label24;
        private System.Windows.Forms.Panel panel6;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.CheckBox checkBox4;
        private System.Windows.Forms.CheckBox cbComments;
        private System.Windows.Forms.CheckBox cbHyperlinks;
        private System.Windows.Forms.CheckBox cbImages;
        private System.Windows.Forms.Panel panel7;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cbExportObject;
        private System.Windows.Forms.Label lblSheetToExport;
        private System.Windows.Forms.ComboBox cbSheet;
        private System.Windows.Forms.CheckBox chPrintHeadings;
        private System.Windows.Forms.CheckBox cbHeadersFooters;
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
            this.panel7 = new System.Windows.Forms.Panel();
            this.cbExportObject = new System.Windows.Forms.ComboBox();
            this.lblSheetToExport = new System.Windows.Forms.Label();
            this.cbSheet = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.panel6 = new System.Windows.Forms.Panel();
            this.cbHeadersFooters = new System.Windows.Forms.CheckBox();
            this.cbImages = new System.Windows.Forms.CheckBox();
            this.cbHyperlinks = new System.Windows.Forms.CheckBox();
            this.cbComments = new System.Windows.Forms.CheckBox();
            this.label6 = new System.Windows.Forms.Label();
            this.panel3 = new System.Windows.Forms.Panel();
            this.edBottom = new System.Windows.Forms.TextBox();
            this.label17 = new System.Windows.Forms.Label();
            this.edRight = new System.Windows.Forms.TextBox();
            this.label16 = new System.Windows.Forms.Label();
            this.edLeft = new System.Windows.Forms.TextBox();
            this.label15 = new System.Windows.Forms.Label();
            this.edTop = new System.Windows.Forms.TextBox();
            this.label14 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.panel8 = new System.Windows.Forms.Panel();
            this.chPrintHeadings = new System.Windows.Forms.CheckBox();
            this.label24 = new System.Windows.Forms.Label();
            this.chFormulaText = new System.Windows.Forms.CheckBox();
            this.chGridLines = new System.Windows.Forms.CheckBox();
            this.checkBox4 = new System.Windows.Forms.CheckBox();
            this.exportDialog = new System.Windows.Forms.SaveFileDialog();
            this.mainToolbar = new System.Windows.Forms.ToolStrip();
            this.openFile = new System.Windows.Forms.ToolStripButton();
            this.export = new System.Windows.Forms.ToolStripButton();
            this.btnClose = new System.Windows.Forms.ToolStripButton();
            this.panel1.SuspendLayout();
            this.panel7.SuspendLayout();
            this.panel6.SuspendLayout();
            this.panel3.SuspendLayout();
            this.panel8.SuspendLayout();
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
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.Controls.Add(this.panel7);
            this.panel1.Controls.Add(this.panel6);
            this.panel1.Controls.Add(this.panel3);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.panel8);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(768, 268);
            this.panel1.TabIndex = 3;
            // 
            // panel7
            // 
            this.panel7.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel7.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.panel7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel7.Controls.Add(this.cbExportObject);
            this.panel7.Controls.Add(this.lblSheetToExport);
            this.panel7.Controls.Add(this.cbSheet);
            this.panel7.Controls.Add(this.label2);
            this.panel7.Location = new System.Drawing.Point(32, 52);
            this.panel7.Name = "panel7";
            this.panel7.Size = new System.Drawing.Size(328, 200);
            this.panel7.TabIndex = 44;
            // 
            // cbExportObject
            // 
            this.cbExportObject.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cbExportObject.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbExportObject.Items.AddRange(new object[] {
            "All sheets",
            "Active Sheet:"});
            this.cbExportObject.Location = new System.Drawing.Point(8, 32);
            this.cbExportObject.Name = "cbExportObject";
            this.cbExportObject.Size = new System.Drawing.Size(293, 21);
            this.cbExportObject.TabIndex = 46;
            this.cbExportObject.SelectedIndexChanged += new System.EventHandler(this.cbExportObject_SelectedIndexChanged);
            // 
            // lblSheetToExport
            // 
            this.lblSheetToExport.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSheetToExport.Location = new System.Drawing.Point(8, 66);
            this.lblSheetToExport.Name = "lblSheetToExport";
            this.lblSheetToExport.Size = new System.Drawing.Size(96, 16);
            this.lblSheetToExport.TabIndex = 45;
            this.lblSheetToExport.Text = "Sheet to export:";
            // 
            // cbSheet
            // 
            this.cbSheet.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cbSheet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbSheet.Location = new System.Drawing.Point(8, 82);
            this.cbSheet.Name = "cbSheet";
            this.cbSheet.Size = new System.Drawing.Size(294, 21);
            this.cbSheet.TabIndex = 44;
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(8, 8);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(192, 16);
            this.label2.TabIndex = 19;
            this.label2.Text = "What to Export:";
            // 
            // panel6
            // 
            this.panel6.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.panel6.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.panel6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel6.Controls.Add(this.cbHeadersFooters);
            this.panel6.Controls.Add(this.cbImages);
            this.panel6.Controls.Add(this.cbHyperlinks);
            this.panel6.Controls.Add(this.cbComments);
            this.panel6.Controls.Add(this.label6);
            this.panel6.Location = new System.Drawing.Point(366, 148);
            this.panel6.Name = "panel6";
            this.panel6.Size = new System.Drawing.Size(176, 104);
            this.panel6.TabIndex = 42;
            // 
            // cbHeadersFooters
            // 
            this.cbHeadersFooters.Location = new System.Drawing.Point(96, 40);
            this.cbHeadersFooters.Name = "cbHeadersFooters";
            this.cbHeadersFooters.Size = new System.Drawing.Size(72, 44);
            this.cbHeadersFooters.TabIndex = 23;
            this.cbHeadersFooters.Text = "Headers / Footers";
            // 
            // cbImages
            // 
            this.cbImages.Checked = true;
            this.cbImages.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbImages.Location = new System.Drawing.Point(16, 32);
            this.cbImages.Name = "cbImages";
            this.cbImages.Size = new System.Drawing.Size(72, 24);
            this.cbImages.TabIndex = 22;
            this.cbImages.Text = "Images";
            // 
            // cbHyperlinks
            // 
            this.cbHyperlinks.Checked = true;
            this.cbHyperlinks.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbHyperlinks.Location = new System.Drawing.Point(16, 80);
            this.cbHyperlinks.Name = "cbHyperlinks";
            this.cbHyperlinks.Size = new System.Drawing.Size(80, 24);
            this.cbHyperlinks.TabIndex = 21;
            this.cbHyperlinks.Text = "HyperLinks";
            // 
            // cbComments
            // 
            this.cbComments.Checked = true;
            this.cbComments.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbComments.Location = new System.Drawing.Point(16, 56);
            this.cbComments.Name = "cbComments";
            this.cbComments.Size = new System.Drawing.Size(80, 24);
            this.cbComments.TabIndex = 20;
            this.cbComments.Text = "Comments";
            // 
            // label6
            // 
            this.label6.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(8, 16);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(192, 16);
            this.label6.TabIndex = 19;
            this.label6.Text = "Objects to Export:";
            // 
            // panel3
            // 
            this.panel3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.panel3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel3.Controls.Add(this.edBottom);
            this.panel3.Controls.Add(this.label17);
            this.panel3.Controls.Add(this.edRight);
            this.panel3.Controls.Add(this.label16);
            this.panel3.Controls.Add(this.edLeft);
            this.panel3.Controls.Add(this.label15);
            this.panel3.Controls.Add(this.edTop);
            this.panel3.Controls.Add(this.label14);
            this.panel3.Controls.Add(this.label13);
            this.panel3.Controls.Add(this.label12);
            this.panel3.Location = new System.Drawing.Point(548, 52);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(208, 200);
            this.panel3.TabIndex = 30;
            // 
            // edBottom
            // 
            this.edBottom.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edBottom.Location = new System.Drawing.Point(80, 136);
            this.edBottom.Name = "edBottom";
            this.edBottom.Size = new System.Drawing.Size(48, 20);
            this.edBottom.TabIndex = 26;
            this.edBottom.Text = "0";
            // 
            // label17
            // 
            this.label17.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label17.Location = new System.Drawing.Point(16, 160);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(56, 16);
            this.label17.TabIndex = 25;
            this.label17.Text = "Last Col:";
            // 
            // edRight
            // 
            this.edRight.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edRight.Location = new System.Drawing.Point(80, 160);
            this.edRight.Name = "edRight";
            this.edRight.Size = new System.Drawing.Size(48, 20);
            this.edRight.TabIndex = 24;
            this.edRight.Text = "0";
            // 
            // label16
            // 
            this.label16.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label16.Location = new System.Drawing.Point(16, 136);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(85, 16);
            this.label16.TabIndex = 23;
            this.label16.Text = "Last Row:";
            // 
            // edLeft
            // 
            this.edLeft.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edLeft.Location = new System.Drawing.Point(80, 112);
            this.edLeft.Name = "edLeft";
            this.edLeft.Size = new System.Drawing.Size(48, 20);
            this.edLeft.TabIndex = 22;
            this.edLeft.Text = "0";
            // 
            // label15
            // 
            this.label15.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label15.Location = new System.Drawing.Point(16, 112);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(85, 16);
            this.label15.TabIndex = 21;
            this.label15.Text = "First Col:";
            // 
            // edTop
            // 
            this.edTop.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edTop.Location = new System.Drawing.Point(80, 88);
            this.edTop.Name = "edTop";
            this.edTop.Size = new System.Drawing.Size(48, 20);
            this.edTop.TabIndex = 20;
            this.edTop.Text = "0";
            // 
            // label14
            // 
            this.label14.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label14.Location = new System.Drawing.Point(16, 88);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(85, 16);
            this.label14.TabIndex = 3;
            this.label14.Text = "First Row:";
            // 
            // label13
            // 
            this.label13.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label13.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label13.Location = new System.Drawing.Point(8, 32);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(184, 32);
            this.label13.TabIndex = 2;
            this.label13.Text = "If any value is <=0 all print_range will be printed";
            // 
            // label12
            // 
            this.label12.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.Location = new System.Drawing.Point(8, 16);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(192, 16);
            this.label12.TabIndex = 1;
            this.label12.Text = "Range to Export:";
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(40, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(80, 16);
            this.label1.TabIndex = 0;
            this.label1.Text = "File to export:";
            // 
            // panel8
            // 
            this.panel8.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.panel8.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.panel8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel8.Controls.Add(this.chPrintHeadings);
            this.panel8.Controls.Add(this.label24);
            this.panel8.Controls.Add(this.chFormulaText);
            this.panel8.Controls.Add(this.chGridLines);
            this.panel8.Location = new System.Drawing.Point(366, 52);
            this.panel8.Name = "panel8";
            this.panel8.Size = new System.Drawing.Size(176, 88);
            this.panel8.TabIndex = 37;
            // 
            // chPrintHeadings
            // 
            this.chPrintHeadings.Location = new System.Drawing.Point(16, 44);
            this.chPrintHeadings.Name = "chPrintHeadings";
            this.chPrintHeadings.Size = new System.Drawing.Size(144, 16);
            this.chPrintHeadings.TabIndex = 20;
            this.chPrintHeadings.Text = "Print Headings";
            // 
            // label24
            // 
            this.label24.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label24.Location = new System.Drawing.Point(8, 8);
            this.label24.Name = "label24";
            this.label24.Size = new System.Drawing.Size(192, 16);
            this.label24.TabIndex = 19;
            this.label24.Text = "Export Options:";
            // 
            // chFormulaText
            // 
            this.chFormulaText.Location = new System.Drawing.Point(16, 64);
            this.chFormulaText.Name = "chFormulaText";
            this.chFormulaText.Size = new System.Drawing.Size(136, 16);
            this.chFormulaText.TabIndex = 17;
            this.chFormulaText.Text = "Print Formula Text";
            // 
            // chGridLines
            // 
            this.chGridLines.Location = new System.Drawing.Point(16, 24);
            this.chGridLines.Name = "chGridLines";
            this.chGridLines.Size = new System.Drawing.Size(128, 16);
            this.chGridLines.TabIndex = 16;
            this.chGridLines.Text = "Print Grid Lines";
            // 
            // checkBox4
            // 
            this.checkBox4.Location = new System.Drawing.Point(0, 0);
            this.checkBox4.Name = "checkBox4";
            this.checkBox4.Size = new System.Drawing.Size(104, 24);
            this.checkBox4.TabIndex = 0;
            // 
            // exportDialog
            // 
            this.exportDialog.DefaultExt = "svg";
            this.exportDialog.Filter = "SVG Files|*.svg";
            this.exportDialog.Title = "Files will be saved as Filename_sheetname_pagenumber.svg";
            // 
            // mainToolbar
            // 
            this.mainToolbar.ImageScalingSize = new System.Drawing.Size(24, 24);
            this.mainToolbar.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openFile,
            this.export,
            this.btnClose});
            this.mainToolbar.Location = new System.Drawing.Point(0, 0);
            this.mainToolbar.Name = "mainToolbar";
            this.mainToolbar.Size = new System.Drawing.Size(768, 31);
            this.mainToolbar.TabIndex = 8;
            // 
            // openFile
            // 
            this.openFile.Image = ((System.Drawing.Image)(resources.GetObject("openFile.Image")));
            this.openFile.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.openFile.Name = "openFile";
            this.openFile.Size = new System.Drawing.Size(85, 28);
            this.openFile.Text = "Open File";
            this.openFile.Click += new System.EventHandler(this.openFile_Click);
            // 
            // export
            // 
            this.export.Image = ((System.Drawing.Image)(resources.GetObject("export.Image")));
            this.export.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.export.Name = "export";
            this.export.Size = new System.Drawing.Size(106, 28);
            this.export.Text = "Export as SVG";
            this.export.Click += new System.EventHandler(this.export_Click);
            // 
            // btnClose
            // 
            this.btnClose.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.btnClose.Image = ((System.Drawing.Image)(resources.GetObject("btnClose.Image")));
            this.btnClose.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(53, 28);
            this.btnClose.Text = "Exit";
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // mainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(768, 268);
            this.Controls.Add(this.mainToolbar);
            this.Controls.Add(this.panel1);
            this.Name = "mainForm";
            this.Text = "Export an Excel file to SVG";
            this.Load += new System.EventHandler(this.mainForm_Load);
            this.panel1.ResumeLayout(false);
            this.panel7.ResumeLayout(false);
            this.panel6.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.panel8.ResumeLayout(false);
            this.mainToolbar.ResumeLayout(false);
            this.mainToolbar.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        private ToolStrip mainToolbar;
        private ToolStripButton openFile;
        private ToolStripButton export;
        private ToolStripButton btnClose;
    }
}

