using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using FlexCel.Render;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.Runtime.InteropServices;
namespace PrintPreviewandExport
{
    public partial class mainForm: System.Windows.Forms.Form
    {
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.PrintDialog printDialog1;
        private System.Windows.Forms.PrintPreviewDialog printPreviewDialog1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox edFileName;
        private System.Windows.Forms.CheckBox chFormulaText;
        private System.Windows.Forms.CheckBox chAntiAlias;
        private System.Windows.Forms.CheckBox chGridLines;
        private System.Windows.Forms.TextBox edHeader;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox edFooter;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox edHPages;
        private System.Windows.Forms.TextBox edVPages;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.CheckBox chPrintLeft;
        private System.Windows.Forms.CheckBox chFitIn;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox edZoom;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox edl;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox edt;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox edr;
        private System.Windows.Forms.Label labelb;
        private System.Windows.Forms.TextBox edb;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.TextBox edh;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox edf;
        private System.Windows.Forms.CheckBox Landscape;
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
        private System.Windows.Forms.ComboBox cbSheet;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.CheckBox cbConfidential;
        private System.Windows.Forms.SaveFileDialog exportImageDialog;
        private System.Windows.Forms.CheckBox chHeadings;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.ComboBox cbInterpolation;
        private System.Windows.Forms.SaveFileDialog exportTiffDialog;
        private System.Windows.Forms.CheckBox cbAllSheets;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.CheckBox cbResetPageNumber;
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
            this.printDialog1 = new System.Windows.Forms.PrintDialog();
            this.flexCelPrintDocument1 = new FlexCel.Render.FlexCelPrintDocument();
            this.printPreviewDialog1 = new System.Windows.Forms.PrintPreviewDialog();
            this.panel1 = new System.Windows.Forms.Panel();
            this.cbResetPageNumber = new System.Windows.Forms.CheckBox();
            this.panel4 = new System.Windows.Forms.Panel();
            this.cbAllSheets = new System.Windows.Forms.CheckBox();
            this.label19 = new System.Windows.Forms.Label();
            this.cbInterpolation = new System.Windows.Forms.ComboBox();
            this.chHeadings = new System.Windows.Forms.CheckBox();
            this.cbConfidential = new System.Windows.Forms.CheckBox();
            this.label18 = new System.Windows.Forms.Label();
            this.cbSheet = new System.Windows.Forms.ComboBox();
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
            this.Landscape = new System.Windows.Forms.CheckBox();
            this.label11 = new System.Windows.Forms.Label();
            this.edf = new System.Windows.Forms.TextBox();
            this.labelb = new System.Windows.Forms.Label();
            this.edb = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.edr = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.edt = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.edl = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.edZoom = new System.Windows.Forms.TextBox();
            this.chFitIn = new System.Windows.Forms.CheckBox();
            this.chPrintLeft = new System.Windows.Forms.CheckBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.edVPages = new System.Windows.Forms.TextBox();
            this.edHPages = new System.Windows.Forms.TextBox();
            this.edFooter = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.edHeader = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.edFileName = new System.Windows.Forms.TextBox();
            this.chFormulaText = new System.Windows.Forms.CheckBox();
            this.chGridLines = new System.Windows.Forms.CheckBox();
            this.chAntiAlias = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.edh = new System.Windows.Forms.TextBox();
            this.exportImageDialog = new System.Windows.Forms.SaveFileDialog();
            this.exportTiffDialog = new System.Windows.Forms.SaveFileDialog();
            this.mainToolbar = new System.Windows.Forms.ToolStrip();
            this.btnOpenFile = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.btnSetup = new System.Windows.Forms.ToolStripButton();
            this.btnPreview = new System.Windows.Forms.ToolStripButton();
            this.btnPrint = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.btnExportAsImages = new System.Windows.Forms.ToolStripDropDownButton();
            this.usingFlexCelImgExportAllPagesrecommendedWayToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.blackAndWhiteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.colorsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.trueColorToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.usingFlexCelImgExport1PagerecommendedWayToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.blackAndWhiteToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.colorsToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.trueColorToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.multiPageTIFFUsingFlexCelImgExportToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.faxToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.blackAndWhiteToolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
            this.colorsToolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
            this.trueColorToolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
            this.btnExit = new System.Windows.Forms.ToolStripButton();
            this.panel1.SuspendLayout();
            this.panel3.SuspendLayout();
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
            // printDialog1
            // 
            this.printDialog1.AllowSomePages = true;
            this.printDialog1.Document = this.flexCelPrintDocument1;
            // 
            // flexCelPrintDocument1
            // 
            this.flexCelPrintDocument1.AllVisibleSheets = false;
            this.flexCelPrintDocument1.ResetPageNumberOnEachSheet = false;
            this.flexCelPrintDocument1.Workbook = null;
            this.flexCelPrintDocument1.GetPrinterHardMargins += new FlexCel.Render.PrintHardMarginsEventHandler(this.flexCelPrintDocument1_GetPrinterHardMargins);
            this.flexCelPrintDocument1.BeforePrintPage += new System.Drawing.Printing.PrintPageEventHandler(this.flexCelPrintDocument1_BeforePrintPage);
            this.flexCelPrintDocument1.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(this.flexCelPrintDocument1_PrintPage);
            // 
            // printPreviewDialog1
            // 
            this.printPreviewDialog1.AutoScrollMargin = new System.Drawing.Size(0, 0);
            this.printPreviewDialog1.AutoScrollMinSize = new System.Drawing.Size(0, 0);
            this.printPreviewDialog1.ClientSize = new System.Drawing.Size(400, 300);
            this.printPreviewDialog1.Document = this.flexCelPrintDocument1;
            this.printPreviewDialog1.Enabled = true;
            this.printPreviewDialog1.Icon = ((System.Drawing.Icon)(resources.GetObject("printPreviewDialog1.Icon")));
            this.printPreviewDialog1.Name = "printPreviewDialog1";
            this.printPreviewDialog1.Visible = false;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.Controls.Add(this.cbResetPageNumber);
            this.panel1.Controls.Add(this.panel4);
            this.panel1.Controls.Add(this.cbAllSheets);
            this.panel1.Controls.Add(this.label19);
            this.panel1.Controls.Add(this.cbInterpolation);
            this.panel1.Controls.Add(this.chHeadings);
            this.panel1.Controls.Add(this.cbConfidential);
            this.panel1.Controls.Add(this.label18);
            this.panel1.Controls.Add(this.cbSheet);
            this.panel1.Controls.Add(this.panel3);
            this.panel1.Controls.Add(this.Landscape);
            this.panel1.Controls.Add(this.label11);
            this.panel1.Controls.Add(this.edf);
            this.panel1.Controls.Add(this.labelb);
            this.panel1.Controls.Add(this.edb);
            this.panel1.Controls.Add(this.label9);
            this.panel1.Controls.Add(this.edr);
            this.panel1.Controls.Add(this.label8);
            this.panel1.Controls.Add(this.edt);
            this.panel1.Controls.Add(this.label7);
            this.panel1.Controls.Add(this.edl);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.edZoom);
            this.panel1.Controls.Add(this.chFitIn);
            this.panel1.Controls.Add(this.chPrintLeft);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.edVPages);
            this.panel1.Controls.Add(this.edHPages);
            this.panel1.Controls.Add(this.edFooter);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.edHeader);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.edFileName);
            this.panel1.Controls.Add(this.chFormulaText);
            this.panel1.Controls.Add(this.chGridLines);
            this.panel1.Controls.Add(this.chAntiAlias);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.label10);
            this.panel1.Controls.Add(this.edh);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 38);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(768, 479);
            this.panel1.TabIndex = 3;
            // 
            // cbResetPageNumber
            // 
            this.cbResetPageNumber.Enabled = false;
            this.cbResetPageNumber.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbResetPageNumber.Location = new System.Drawing.Point(528, 48);
            this.cbResetPageNumber.Name = "cbResetPageNumber";
            this.cbResetPageNumber.Size = new System.Drawing.Size(216, 16);
            this.cbResetPageNumber.TabIndex = 39;
            this.cbResetPageNumber.Text = "Reset Page number on each sheet.";
            // 
            // panel4
            // 
            this.panel4.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel4.Location = new System.Drawing.Point(16, 72);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(736, 3);
            this.panel4.TabIndex = 38;
            // 
            // cbAllSheets
            // 
            this.cbAllSheets.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbAllSheets.Location = new System.Drawing.Point(32, 48);
            this.cbAllSheets.Name = "cbAllSheets";
            this.cbAllSheets.Size = new System.Drawing.Size(104, 16);
            this.cbAllSheets.TabIndex = 37;
            this.cbAllSheets.Text = "All Sheets";
            this.cbAllSheets.CheckedChanged += new System.EventHandler(this.cbAllSheets_CheckedChanged);
            // 
            // label19
            // 
            this.label19.Location = new System.Drawing.Point(392, 80);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(160, 40);
            this.label19.TabIndex = 36;
            this.label19.Text = "Interpolation mode for images: Sometimes a lower mode might give crisper results." +
    "";
            // 
            // cbInterpolation
            // 
            this.cbInterpolation.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbInterpolation.Items.AddRange(new object[] {
            "Bicubic",
            "Bilinear",
            "Default",
            "High",
            "HighQualityBicubic",
            "HighQualityBilinear ",
            "Low",
            "NearestNeighbor"});
            this.cbInterpolation.Location = new System.Drawing.Point(560, 88);
            this.cbInterpolation.Name = "cbInterpolation";
            this.cbInterpolation.Size = new System.Drawing.Size(152, 21);
            this.cbInterpolation.TabIndex = 35;
            // 
            // chHeadings
            // 
            this.chHeadings.Location = new System.Drawing.Point(176, 136);
            this.chHeadings.Name = "chHeadings";
            this.chHeadings.Size = new System.Drawing.Size(128, 24);
            this.chHeadings.TabIndex = 34;
            this.chHeadings.Text = "Print Headings";
            // 
            // cbConfidential
            // 
            this.cbConfidential.Location = new System.Drawing.Point(56, 112);
            this.cbConfidential.Name = "cbConfidential";
            this.cbConfidential.Size = new System.Drawing.Size(232, 16);
            this.cbConfidential.TabIndex = 33;
            this.cbConfidential.Text = "Print \"Confidential\" on each page";
            // 
            // label18
            // 
            this.label18.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label18.Location = new System.Drawing.Point(168, 48);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(88, 16);
            this.label18.TabIndex = 32;
            this.label18.Text = "Sheet to print:";
            // 
            // cbSheet
            // 
            this.cbSheet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbSheet.Location = new System.Drawing.Point(256, 43);
            this.cbSheet.Name = "cbSheet";
            this.cbSheet.Size = new System.Drawing.Size(160, 21);
            this.cbSheet.TabIndex = 31;
            this.cbSheet.SelectedIndexChanged += new System.EventHandler(this.cbSheet_SelectedIndexChanged);
            // 
            // panel3
            // 
            this.panel3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
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
            this.panel3.Location = new System.Drawing.Point(504, 232);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(216, 224);
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
            this.label14.Location = new System.Drawing.Point(8, 88);
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
            this.label13.Size = new System.Drawing.Size(192, 32);
            this.label13.TabIndex = 2;
            this.label13.Text = "If one of this values is <=0 all print_range will be printed";
            // 
            // label12
            // 
            this.label12.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.Location = new System.Drawing.Point(8, 16);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(192, 16);
            this.label12.TabIndex = 1;
            this.label12.Text = "Range to Print:";
            // 
            // Landscape
            // 
            this.Landscape.Location = new System.Drawing.Point(456, 136);
            this.Landscape.Name = "Landscape";
            this.Landscape.Size = new System.Drawing.Size(96, 24);
            this.Landscape.TabIndex = 29;
            this.Landscape.Text = "Landscape";
            // 
            // label11
            // 
            this.label11.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.Location = new System.Drawing.Point(264, 416);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(80, 16);
            this.label11.TabIndex = 28;
            this.label11.Text = "Footer Margin";
            // 
            // edf
            // 
            this.edf.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edf.Location = new System.Drawing.Point(344, 416);
            this.edf.Name = "edf";
            this.edf.Size = new System.Drawing.Size(128, 20);
            this.edf.TabIndex = 27;
            // 
            // labelb
            // 
            this.labelb.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelb.Location = new System.Drawing.Point(256, 368);
            this.labelb.Name = "labelb";
            this.labelb.Size = new System.Drawing.Size(88, 16);
            this.labelb.TabIndex = 26;
            this.labelb.Text = "Bottom Margin";
            // 
            // edb
            // 
            this.edb.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edb.Location = new System.Drawing.Point(344, 368);
            this.edb.Name = "edb";
            this.edb.Size = new System.Drawing.Size(128, 20);
            this.edb.TabIndex = 25;
            // 
            // label9
            // 
            this.label9.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.Location = new System.Drawing.Point(56, 368);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(80, 16);
            this.label9.TabIndex = 24;
            this.label9.Text = "Right Margin";
            // 
            // edr
            // 
            this.edr.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edr.Location = new System.Drawing.Point(136, 368);
            this.edr.Name = "edr";
            this.edr.Size = new System.Drawing.Size(112, 20);
            this.edr.TabIndex = 23;
            // 
            // label8
            // 
            this.label8.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(264, 328);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(80, 16);
            this.label8.TabIndex = 22;
            this.label8.Text = "Top Margin";
            // 
            // edt
            // 
            this.edt.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edt.Location = new System.Drawing.Point(344, 328);
            this.edt.Name = "edt";
            this.edt.Size = new System.Drawing.Size(128, 20);
            this.edt.TabIndex = 21;
            // 
            // label7
            // 
            this.label7.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(56, 328);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(80, 16);
            this.label7.TabIndex = 20;
            this.label7.Text = "Left Margin";
            // 
            // edl
            // 
            this.edl.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edl.Location = new System.Drawing.Point(136, 328);
            this.edl.Name = "edl";
            this.edl.Size = new System.Drawing.Size(112, 20);
            this.edl.TabIndex = 19;
            // 
            // label4
            // 
            this.label4.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(120, 280);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(56, 16);
            this.label4.TabIndex = 18;
            this.label4.Text = "Zoom (%)";
            // 
            // edZoom
            // 
            this.edZoom.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edZoom.Location = new System.Drawing.Point(184, 280);
            this.edZoom.Name = "edZoom";
            this.edZoom.Size = new System.Drawing.Size(24, 20);
            this.edZoom.TabIndex = 17;
            // 
            // chFitIn
            // 
            this.chFitIn.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chFitIn.Location = new System.Drawing.Point(56, 248);
            this.chFitIn.Name = "chFitIn";
            this.chFitIn.Size = new System.Drawing.Size(56, 24);
            this.chFitIn.TabIndex = 16;
            this.chFitIn.Text = "Fit in";
            this.chFitIn.CheckedChanged += new System.EventHandler(this.chFitIn_CheckedChanged);
            // 
            // chPrintLeft
            // 
            this.chPrintLeft.Location = new System.Drawing.Point(312, 136);
            this.chPrintLeft.Name = "chPrintLeft";
            this.chPrintLeft.Size = new System.Drawing.Size(136, 24);
            this.chPrintLeft.TabIndex = 15;
            this.chPrintLeft.Text = "Print Left, then down.";
            // 
            // label6
            // 
            this.label6.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(256, 248);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(80, 16);
            this.label6.TabIndex = 14;
            this.label6.Text = "pages tall.";
            // 
            // label5
            // 
            this.label5.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(144, 248);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(80, 16);
            this.label5.TabIndex = 13;
            this.label5.Text = "pages wide x";
            // 
            // edVPages
            // 
            this.edVPages.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edVPages.Location = new System.Drawing.Point(224, 248);
            this.edVPages.Name = "edVPages";
            this.edVPages.ReadOnly = true;
            this.edVPages.Size = new System.Drawing.Size(24, 20);
            this.edVPages.TabIndex = 12;
            // 
            // edHPages
            // 
            this.edHPages.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edHPages.Location = new System.Drawing.Point(112, 248);
            this.edHPages.Name = "edHPages";
            this.edHPages.ReadOnly = true;
            this.edHPages.Size = new System.Drawing.Size(24, 20);
            this.edHPages.TabIndex = 10;
            // 
            // edFooter
            // 
            this.edFooter.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.edFooter.BackColor = System.Drawing.Color.White;
            this.edFooter.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edFooter.Location = new System.Drawing.Point(112, 200);
            this.edFooter.Name = "edFooter";
            this.edFooter.Size = new System.Drawing.Size(608, 20);
            this.edFooter.TabIndex = 8;
            // 
            // label3
            // 
            this.label3.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(56, 200);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(56, 16);
            this.label3.TabIndex = 7;
            this.label3.Text = "Footer:";
            // 
            // edHeader
            // 
            this.edHeader.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.edHeader.BackColor = System.Drawing.Color.White;
            this.edHeader.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edHeader.Location = new System.Drawing.Point(112, 176);
            this.edHeader.Name = "edHeader";
            this.edHeader.Size = new System.Drawing.Size(608, 20);
            this.edHeader.TabIndex = 6;
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(56, 176);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(56, 16);
            this.label2.TabIndex = 5;
            this.label2.Text = "Header:";
            // 
            // edFileName
            // 
            this.edFileName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.edFileName.BackColor = System.Drawing.Color.White;
            this.edFileName.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.edFileName.Location = new System.Drawing.Point(112, 16);
            this.edFileName.Name = "edFileName";
            this.edFileName.ReadOnly = true;
            this.edFileName.Size = new System.Drawing.Size(632, 13);
            this.edFileName.TabIndex = 4;
            this.edFileName.Text = "No file selected";
            // 
            // chFormulaText
            // 
            this.chFormulaText.Location = new System.Drawing.Point(576, 136);
            this.chFormulaText.Name = "chFormulaText";
            this.chFormulaText.Size = new System.Drawing.Size(136, 24);
            this.chFormulaText.TabIndex = 3;
            this.chFormulaText.Text = "Print Formula Text";
            // 
            // chGridLines
            // 
            this.chGridLines.Location = new System.Drawing.Point(56, 136);
            this.chGridLines.Name = "chGridLines";
            this.chGridLines.Size = new System.Drawing.Size(104, 24);
            this.chGridLines.TabIndex = 2;
            this.chGridLines.Text = "Print Grid Lines";
            // 
            // chAntiAlias
            // 
            this.chAntiAlias.Location = new System.Drawing.Point(56, 88);
            this.chAntiAlias.Name = "chAntiAlias";
            this.chAntiAlias.Size = new System.Drawing.Size(152, 16);
            this.chAntiAlias.TabIndex = 1;
            this.chAntiAlias.Text = "Antialias Text";
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(24, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(80, 16);
            this.label1.TabIndex = 0;
            this.label1.Text = "File to print:";
            // 
            // label10
            // 
            this.label10.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.Location = new System.Drawing.Point(48, 416);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(88, 16);
            this.label10.TabIndex = 22;
            this.label10.Text = "Header Margin";
            // 
            // edh
            // 
            this.edh.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edh.Location = new System.Drawing.Point(136, 416);
            this.edh.Name = "edh";
            this.edh.Size = new System.Drawing.Size(112, 20);
            this.edh.TabIndex = 21;
            // 
            // exportImageDialog
            // 
            this.exportImageDialog.DefaultExt = "png";
            this.exportImageDialog.Filter = "Png files|*.png|Jpg files|*.jpg";
            this.exportImageDialog.Title = "Save image as...";
            // 
            // exportTiffDialog
            // 
            this.exportTiffDialog.DefaultExt = "tif";
            this.exportTiffDialog.Filter = "TIFF Files|*.tif";
            this.exportTiffDialog.Title = "Save image as multi page tiff...";
            // 
            // mainToolbar
            // 
            this.mainToolbar.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnOpenFile,
            this.toolStripSeparator1,
            this.btnSetup,
            this.btnPreview,
            this.btnPrint,
            this.toolStripSeparator2,
            this.btnExportAsImages,
            this.btnExit});
            this.mainToolbar.Location = new System.Drawing.Point(0, 0);
            this.mainToolbar.Name = "mainToolbar";
            this.mainToolbar.Size = new System.Drawing.Size(768, 38);
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
            this.btnOpenFile.Click += new System.EventHandler(this.openFile_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 46);
            // 
            // btnSetup
            // 
            this.btnSetup.Image = ((System.Drawing.Image)(resources.GetObject("btnSetup.Image")));
            this.btnSetup.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnSetup.Name = "btnSetup";
            this.btnSetup.Size = new System.Drawing.Size(69, 35);
            this.btnSetup.Text = "Print &Setup";
            this.btnSetup.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnSetup.Click += new System.EventHandler(this.setup_Click);
            // 
            // btnPreview
            // 
            this.btnPreview.Image = ((System.Drawing.Image)(resources.GetObject("btnPreview.Image")));
            this.btnPreview.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnPreview.Name = "btnPreview";
            this.btnPreview.Size = new System.Drawing.Size(80, 35);
            this.btnPreview.Text = "Print Pre&view";
            this.btnPreview.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnPreview.Click += new System.EventHandler(this.preview_Click);
            // 
            // btnPrint
            // 
            this.btnPrint.Image = ((System.Drawing.Image)(resources.GetObject("btnPrint.Image")));
            this.btnPrint.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(36, 35);
            this.btnPrint.Text = "&Print";
            this.btnPrint.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnPrint.Click += new System.EventHandler(this.print_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(6, 46);
            // 
            // btnExportAsImages
            // 
            this.btnExportAsImages.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.usingFlexCelImgExportAllPagesrecommendedWayToolStripMenuItem,
            this.usingFlexCelImgExport1PagerecommendedWayToolStripMenuItem,
            this.multiPageTIFFUsingFlexCelImgExportToolStripMenuItem});
            this.btnExportAsImages.Image = ((System.Drawing.Image)(resources.GetObject("btnExportAsImages.Image")));
            this.btnExportAsImages.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnExportAsImages.Name = "btnExportAsImages";
            this.btnExportAsImages.Size = new System.Drawing.Size(108, 35);
            this.btnExportAsImages.Text = "Export as &Images";
            this.btnExportAsImages.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            // 
            // usingFlexCelImgExportAllPagesrecommendedWayToolStripMenuItem
            // 
            this.usingFlexCelImgExportAllPagesrecommendedWayToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.blackAndWhiteToolStripMenuItem,
            this.colorsToolStripMenuItem,
            this.trueColorToolStripMenuItem});
            this.usingFlexCelImgExportAllPagesrecommendedWayToolStripMenuItem.Name = "usingFlexCelImgExportAllPagesrecommendedWayToolStripMenuItem";
            this.usingFlexCelImgExportAllPagesrecommendedWayToolStripMenuItem.Size = new System.Drawing.Size(153, 22);
            this.usingFlexCelImgExportAllPagesrecommendedWayToolStripMenuItem.Text = "All pages";
            // 
            // blackAndWhiteToolStripMenuItem
            // 
            this.blackAndWhiteToolStripMenuItem.Name = "blackAndWhiteToolStripMenuItem";
            this.blackAndWhiteToolStripMenuItem.Size = new System.Drawing.Size(161, 22);
            this.blackAndWhiteToolStripMenuItem.Text = "Black And White";
            this.blackAndWhiteToolStripMenuItem.Click += new System.EventHandler(this.ImgBlackAndWhite_Click);
            // 
            // colorsToolStripMenuItem
            // 
            this.colorsToolStripMenuItem.Name = "colorsToolStripMenuItem";
            this.colorsToolStripMenuItem.Size = new System.Drawing.Size(161, 22);
            this.colorsToolStripMenuItem.Text = "256 Colors";
            this.colorsToolStripMenuItem.Click += new System.EventHandler(this.Img256Colors_Click);
            // 
            // trueColorToolStripMenuItem
            // 
            this.trueColorToolStripMenuItem.Name = "trueColorToolStripMenuItem";
            this.trueColorToolStripMenuItem.Size = new System.Drawing.Size(161, 22);
            this.trueColorToolStripMenuItem.Text = "True Color";
            this.trueColorToolStripMenuItem.Click += new System.EventHandler(this.ImgTrueColor_Click);
            // 
            // usingFlexCelImgExport1PagerecommendedWayToolStripMenuItem
            // 
            this.usingFlexCelImgExport1PagerecommendedWayToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.blackAndWhiteToolStripMenuItem1,
            this.colorsToolStripMenuItem1,
            this.trueColorToolStripMenuItem1});
            this.usingFlexCelImgExport1PagerecommendedWayToolStripMenuItem.Name = "usingFlexCelImgExport1PagerecommendedWayToolStripMenuItem";
            this.usingFlexCelImgExport1PagerecommendedWayToolStripMenuItem.Size = new System.Drawing.Size(153, 22);
            this.usingFlexCelImgExport1PagerecommendedWayToolStripMenuItem.Text = "1 page";
            // 
            // blackAndWhiteToolStripMenuItem1
            // 
            this.blackAndWhiteToolStripMenuItem1.Name = "blackAndWhiteToolStripMenuItem1";
            this.blackAndWhiteToolStripMenuItem1.Size = new System.Drawing.Size(161, 22);
            this.blackAndWhiteToolStripMenuItem1.Text = "Black And White";
            this.blackAndWhiteToolStripMenuItem1.Click += new System.EventHandler(this.ImgBlackAndWhite2_Click);
            // 
            // colorsToolStripMenuItem1
            // 
            this.colorsToolStripMenuItem1.Name = "colorsToolStripMenuItem1";
            this.colorsToolStripMenuItem1.Size = new System.Drawing.Size(161, 22);
            this.colorsToolStripMenuItem1.Text = "256 Colors";
            this.colorsToolStripMenuItem1.Click += new System.EventHandler(this.Img256Colors2_Click);
            // 
            // trueColorToolStripMenuItem1
            // 
            this.trueColorToolStripMenuItem1.Name = "trueColorToolStripMenuItem1";
            this.trueColorToolStripMenuItem1.Size = new System.Drawing.Size(161, 22);
            this.trueColorToolStripMenuItem1.Text = "True Color";
            this.trueColorToolStripMenuItem1.Click += new System.EventHandler(this.ImgTrueColor2_Click);
            // 
            // multiPageTIFFUsingFlexCelImgExportToolStripMenuItem
            // 
            this.multiPageTIFFUsingFlexCelImgExportToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.faxToolStripMenuItem,
            this.blackAndWhiteToolStripMenuItem2,
            this.colorsToolStripMenuItem2,
            this.trueColorToolStripMenuItem2});
            this.multiPageTIFFUsingFlexCelImgExportToolStripMenuItem.Name = "multiPageTIFFUsingFlexCelImgExportToolStripMenuItem";
            this.multiPageTIFFUsingFlexCelImgExportToolStripMenuItem.Size = new System.Drawing.Size(153, 22);
            this.multiPageTIFFUsingFlexCelImgExportToolStripMenuItem.Text = "MultiPage TIFF";
            // 
            // faxToolStripMenuItem
            // 
            this.faxToolStripMenuItem.Name = "faxToolStripMenuItem";
            this.faxToolStripMenuItem.Size = new System.Drawing.Size(161, 22);
            this.faxToolStripMenuItem.Text = "Fax";
            this.faxToolStripMenuItem.Click += new System.EventHandler(this.TiffFax_Click);
            // 
            // blackAndWhiteToolStripMenuItem2
            // 
            this.blackAndWhiteToolStripMenuItem2.Name = "blackAndWhiteToolStripMenuItem2";
            this.blackAndWhiteToolStripMenuItem2.Size = new System.Drawing.Size(161, 22);
            this.blackAndWhiteToolStripMenuItem2.Text = "Black And White";
            this.blackAndWhiteToolStripMenuItem2.Click += new System.EventHandler(this.TiffBlackAndWhite_Click);
            // 
            // colorsToolStripMenuItem2
            // 
            this.colorsToolStripMenuItem2.Name = "colorsToolStripMenuItem2";
            this.colorsToolStripMenuItem2.Size = new System.Drawing.Size(161, 22);
            this.colorsToolStripMenuItem2.Text = "256 Colors";
            this.colorsToolStripMenuItem2.Click += new System.EventHandler(this.Tiff256Colors_Click);
            // 
            // trueColorToolStripMenuItem2
            // 
            this.trueColorToolStripMenuItem2.Name = "trueColorToolStripMenuItem2";
            this.trueColorToolStripMenuItem2.Size = new System.Drawing.Size(161, 22);
            this.trueColorToolStripMenuItem2.Text = "True Color";
            this.trueColorToolStripMenuItem2.Click += new System.EventHandler(this.TiffTrueColor_Click);
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
            this.ClientSize = new System.Drawing.Size(768, 517);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.mainToolbar);
            this.Name = "mainForm";
            this.Text = "Print and preview a file";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.mainToolbar.ResumeLayout(false);
            this.mainToolbar.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        private ToolStrip mainToolbar;
        private ToolStripButton btnOpenFile;
        private ToolStripSeparator toolStripSeparator1;
        private ToolStripButton btnSetup;
        private ToolStripButton btnPreview;
        private ToolStripButton btnPrint;
        private ToolStripSeparator toolStripSeparator2;
        private ToolStripButton btnExit;
        private ToolStripDropDownButton btnExportAsImages;
        private ToolStripMenuItem usingFlexCelImgExportAllPagesrecommendedWayToolStripMenuItem;
        private ToolStripMenuItem blackAndWhiteToolStripMenuItem;
        private ToolStripMenuItem colorsToolStripMenuItem;
        private ToolStripMenuItem trueColorToolStripMenuItem;
        private ToolStripMenuItem usingFlexCelImgExport1PagerecommendedWayToolStripMenuItem;
        private ToolStripMenuItem blackAndWhiteToolStripMenuItem1;
        private ToolStripMenuItem colorsToolStripMenuItem1;
        private ToolStripMenuItem trueColorToolStripMenuItem1;
        private ToolStripMenuItem multiPageTIFFUsingFlexCelImgExportToolStripMenuItem;
        private ToolStripMenuItem faxToolStripMenuItem;
        private ToolStripMenuItem blackAndWhiteToolStripMenuItem2;
        private ToolStripMenuItem colorsToolStripMenuItem2;
        private ToolStripMenuItem trueColorToolStripMenuItem2;
    }
}

