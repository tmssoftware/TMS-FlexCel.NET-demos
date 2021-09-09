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
using System.Drawing.Drawing2D;
using FlexCel.Pdf;
using System.Runtime.InteropServices;
namespace ExportPdf
{
    public partial class mainForm: System.Windows.Forms.Form
    {
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox edFileName;
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
        private System.Windows.Forms.CheckBox chExportAll;
        private System.Windows.Forms.SaveFileDialog exportDialog;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.CheckBox chEmbed;
        private System.Windows.Forms.Label label20;
        private System.Windows.Forms.ComboBox cbFontMapping;
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox edZoom;
        private System.Windows.Forms.CheckBox chFitIn;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox edVPages;
        private System.Windows.Forms.TextBox edHPages;
        private System.Windows.Forms.Label label21;
        private System.Windows.Forms.Panel panel6;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox edf;
        private System.Windows.Forms.Label labelb;
        private System.Windows.Forms.TextBox edb;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox edr;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox edt;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox edl;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.TextBox edh;
        private System.Windows.Forms.Label label22;
        private System.Windows.Forms.Panel panel7;
        private System.Windows.Forms.Label label23;
        private System.Windows.Forms.TextBox edFooter;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox edHeader;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Panel panel8;
        private System.Windows.Forms.CheckBox chPrintLeft;
        private System.Windows.Forms.CheckBox chFormulaText;
        private System.Windows.Forms.CheckBox chGridLines;
        private System.Windows.Forms.Label label24;
        private System.Windows.Forms.Panel panel9;
        private System.Windows.Forms.Label label25;
        private System.Windows.Forms.Label label26;
        private System.Windows.Forms.TextBox edAuthor;
        private System.Windows.Forms.Label label27;
        private System.Windows.Forms.Label label28;
        private System.Windows.Forms.TextBox edTitle;
        private System.Windows.Forms.TextBox edSubject;
        private System.Windows.Forms.CheckBox cbKerning;
        private System.Windows.Forms.CheckBox chLandscape;
        private System.Windows.Forms.CheckBox cbResetPageNumber;
        private System.Windows.Forms.CheckBox cbUseGetFontData;
        private System.Windows.Forms.CheckBox cbConfidential;
        private System.Windows.Forms.CheckBox chSubset;
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
            FlexCel.Pdf.TPdfProperties tPdfProperties1 = new FlexCel.Pdf.TPdfProperties();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.label32 = new System.Windows.Forms.Label();
            this.label31 = new System.Windows.Forms.Label();
            this.label30 = new System.Windows.Forms.Label();
            this.cbTagged = new System.Windows.Forms.ComboBox();
            this.cbVersion = new System.Windows.Forms.ComboBox();
            this.cbPdfType = new System.Windows.Forms.ComboBox();
            this.label34 = new System.Windows.Forms.Label();
            this.cbConfidential = new System.Windows.Forms.CheckBox();
            this.cbUseGetFontData = new System.Windows.Forms.CheckBox();
            this.cbResetPageNumber = new System.Windows.Forms.CheckBox();
            this.panel9 = new System.Windows.Forms.Panel();
            this.label29 = new System.Windows.Forms.Label();
            this.edLang = new System.Windows.Forms.TextBox();
            this.edSubject = new System.Windows.Forms.TextBox();
            this.label28 = new System.Windows.Forms.Label();
            this.edTitle = new System.Windows.Forms.TextBox();
            this.label27 = new System.Windows.Forms.Label();
            this.label26 = new System.Windows.Forms.Label();
            this.edAuthor = new System.Windows.Forms.TextBox();
            this.label25 = new System.Windows.Forms.Label();
            this.edFileName = new System.Windows.Forms.TextBox();
            this.panel8 = new System.Windows.Forms.Panel();
            this.chLandscape = new System.Windows.Forms.CheckBox();
            this.label24 = new System.Windows.Forms.Label();
            this.chPrintLeft = new System.Windows.Forms.CheckBox();
            this.chFormulaText = new System.Windows.Forms.CheckBox();
            this.chGridLines = new System.Windows.Forms.CheckBox();
            this.panel7 = new System.Windows.Forms.Panel();
            this.edFooter = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.edHeader = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label23 = new System.Windows.Forms.Label();
            this.panel6 = new System.Windows.Forms.Panel();
            this.label22 = new System.Windows.Forms.Label();
            this.edf = new System.Windows.Forms.TextBox();
            this.edb = new System.Windows.Forms.TextBox();
            this.edr = new System.Windows.Forms.TextBox();
            this.edt = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.edl = new System.Windows.Forms.TextBox();
            this.edh = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.labelb = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.panel5 = new System.Windows.Forms.Panel();
            this.label21 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.edZoom = new System.Windows.Forms.TextBox();
            this.chFitIn = new System.Windows.Forms.CheckBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.edVPages = new System.Windows.Forms.TextBox();
            this.edHPages = new System.Windows.Forms.TextBox();
            this.panel4 = new System.Windows.Forms.Panel();
            this.chSubset = new System.Windows.Forms.CheckBox();
            this.cbKerning = new System.Windows.Forms.CheckBox();
            this.label20 = new System.Windows.Forms.Label();
            this.cbFontMapping = new System.Windows.Forms.ComboBox();
            this.chEmbed = new System.Windows.Forms.CheckBox();
            this.label19 = new System.Windows.Forms.Label();
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
            this.chExportAll = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.exportDialog = new System.Windows.Forms.SaveFileDialog();
            this.mainToolbar = new System.Windows.Forms.ToolStrip();
            this.openFile = new System.Windows.Forms.ToolStripButton();
            this.export = new System.Windows.Forms.ToolStripButton();
            this.btnClose = new System.Windows.Forms.ToolStripButton();
            this.flexCelPdfExport1 = new FlexCel.Render.FlexCelPdfExport();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel9.SuspendLayout();
            this.panel8.SuspendLayout();
            this.panel7.SuspendLayout();
            this.panel6.SuspendLayout();
            this.panel5.SuspendLayout();
            this.panel4.SuspendLayout();
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
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.Controls.Add(this.panel2);
            this.panel1.Controls.Add(this.cbConfidential);
            this.panel1.Controls.Add(this.cbUseGetFontData);
            this.panel1.Controls.Add(this.cbResetPageNumber);
            this.panel1.Controls.Add(this.panel9);
            this.panel1.Controls.Add(this.edFileName);
            this.panel1.Controls.Add(this.panel8);
            this.panel1.Controls.Add(this.panel7);
            this.panel1.Controls.Add(this.panel6);
            this.panel1.Controls.Add(this.panel5);
            this.panel1.Controls.Add(this.panel4);
            this.panel1.Controls.Add(this.label18);
            this.panel1.Controls.Add(this.cbSheet);
            this.panel1.Controls.Add(this.panel3);
            this.panel1.Controls.Add(this.chExportAll);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 38);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(768, 593);
            this.panel1.TabIndex = 3;
            // 
            // panel2
            // 
            this.panel2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel2.Controls.Add(this.label32);
            this.panel2.Controls.Add(this.label31);
            this.panel2.Controls.Add(this.label30);
            this.panel2.Controls.Add(this.cbTagged);
            this.panel2.Controls.Add(this.cbVersion);
            this.panel2.Controls.Add(this.cbPdfType);
            this.panel2.Controls.Add(this.label34);
            this.panel2.Location = new System.Drawing.Point(32, 146);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(688, 53);
            this.panel2.TabIndex = 42;
            // 
            // label32
            // 
            this.label32.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label32.Location = new System.Drawing.Point(8, 22);
            this.label32.Name = "label32";
            this.label32.Size = new System.Drawing.Size(41, 16);
            this.label32.TabIndex = 41;
            this.label32.Text = "Type:";
            // 
            // label31
            // 
            this.label31.AutoSize = true;
            this.label31.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label31.Location = new System.Drawing.Point(219, 22);
            this.label31.Name = "label31";
            this.label31.Size = new System.Drawing.Size(53, 14);
            this.label31.TabIndex = 40;
            this.label31.Text = "Version:";
            // 
            // label30
            // 
            this.label30.AutoSize = true;
            this.label30.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label30.Location = new System.Drawing.Point(467, 22);
            this.label30.Name = "label30";
            this.label30.Size = new System.Drawing.Size(50, 14);
            this.label30.TabIndex = 39;
            this.label30.Text = "Tagged:";
            // 
            // cbTagged
            // 
            this.cbTagged.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbTagged.Items.AddRange(new object[] {
            "Full",
            "None"});
            this.cbTagged.Location = new System.Drawing.Point(523, 19);
            this.cbTagged.Name = "cbTagged";
            this.cbTagged.Size = new System.Drawing.Size(149, 21);
            this.cbTagged.TabIndex = 36;
            // 
            // cbVersion
            // 
            this.cbVersion.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbVersion.Items.AddRange(new object[] {
            "1.4 (Acrobat 5)",
            "1.6 (Acrobat 7)"});
            this.cbVersion.Location = new System.Drawing.Point(278, 19);
            this.cbVersion.Name = "cbVersion";
            this.cbVersion.Size = new System.Drawing.Size(168, 21);
            this.cbVersion.TabIndex = 35;
            // 
            // cbPdfType
            // 
            this.cbPdfType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbPdfType.Items.AddRange(new object[] {
            "Standard",
            "PDF/A1",
            "PDF/A2",
            "PDF/A3"});
            this.cbPdfType.Location = new System.Drawing.Point(56, 19);
            this.cbPdfType.Name = "cbPdfType";
            this.cbPdfType.Size = new System.Drawing.Size(144, 21);
            this.cbPdfType.TabIndex = 34;
            // 
            // label34
            // 
            this.label34.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label34.Location = new System.Drawing.Point(8, 0);
            this.label34.Name = "label34";
            this.label34.Size = new System.Drawing.Size(192, 16);
            this.label34.TabIndex = 20;
            this.label34.Text = "Pdf options:";
            // 
            // cbConfidential
            // 
            this.cbConfidential.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.cbConfidential.Location = new System.Drawing.Point(488, 557);
            this.cbConfidential.Name = "cbConfidential";
            this.cbConfidential.Size = new System.Drawing.Size(232, 16);
            this.cbConfidential.TabIndex = 41;
            this.cbConfidential.Text = "Print \"Confidential\" on each page";
            // 
            // cbUseGetFontData
            // 
            this.cbUseGetFontData.Location = new System.Drawing.Point(32, 557);
            this.cbUseGetFontData.Name = "cbUseGetFontData";
            this.cbUseGetFontData.Size = new System.Drawing.Size(312, 16);
            this.cbUseGetFontData.TabIndex = 40;
            this.cbUseGetFontData.Text = "Use UNMANAGED calls to Win32 API to find the fonts.";
            // 
            // cbResetPageNumber
            // 
            this.cbResetPageNumber.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.cbResetPageNumber.Location = new System.Drawing.Point(528, 40);
            this.cbResetPageNumber.Name = "cbResetPageNumber";
            this.cbResetPageNumber.Size = new System.Drawing.Size(200, 16);
            this.cbResetPageNumber.TabIndex = 39;
            this.cbResetPageNumber.Text = "Reset Page number on each sheet";
            // 
            // panel9
            // 
            this.panel9.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel9.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.panel9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel9.Controls.Add(this.label29);
            this.panel9.Controls.Add(this.edLang);
            this.panel9.Controls.Add(this.edSubject);
            this.panel9.Controls.Add(this.label28);
            this.panel9.Controls.Add(this.edTitle);
            this.panel9.Controls.Add(this.label27);
            this.panel9.Controls.Add(this.label26);
            this.panel9.Controls.Add(this.edAuthor);
            this.panel9.Controls.Add(this.label25);
            this.panel9.Location = new System.Drawing.Point(32, 64);
            this.panel9.Name = "panel9";
            this.panel9.Size = new System.Drawing.Size(688, 76);
            this.panel9.TabIndex = 38;
            // 
            // label29
            // 
            this.label29.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label29.Location = new System.Drawing.Point(8, 49);
            this.label29.Name = "label29";
            this.label29.Size = new System.Drawing.Size(48, 16);
            this.label29.TabIndex = 38;
            this.label29.Text = "Lang:";
            // 
            // edLang
            // 
            this.edLang.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edLang.Location = new System.Drawing.Point(56, 46);
            this.edLang.Name = "edLang";
            this.edLang.Size = new System.Drawing.Size(144, 20);
            this.edLang.TabIndex = 37;
            this.edLang.Text = "en-US";
            // 
            // edSubject
            // 
            this.edSubject.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.edSubject.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edSubject.Location = new System.Drawing.Point(278, 45);
            this.edSubject.Name = "edSubject";
            this.edSubject.Size = new System.Drawing.Size(394, 20);
            this.edSubject.TabIndex = 35;
            // 
            // label28
            // 
            this.label28.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label28.Location = new System.Drawing.Point(216, 50);
            this.label28.Name = "label28";
            this.label28.Size = new System.Drawing.Size(56, 16);
            this.label28.TabIndex = 36;
            this.label28.Text = "Subject:";
            // 
            // edTitle
            // 
            this.edTitle.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.edTitle.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edTitle.Location = new System.Drawing.Point(278, 19);
            this.edTitle.Name = "edTitle";
            this.edTitle.Size = new System.Drawing.Size(394, 20);
            this.edTitle.TabIndex = 33;
            // 
            // label27
            // 
            this.label27.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label27.Location = new System.Drawing.Point(232, 22);
            this.label27.Name = "label27";
            this.label27.Size = new System.Drawing.Size(48, 16);
            this.label27.TabIndex = 34;
            this.label27.Text = "Title:";
            // 
            // label26
            // 
            this.label26.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label26.Location = new System.Drawing.Point(8, 23);
            this.label26.Name = "label26";
            this.label26.Size = new System.Drawing.Size(48, 16);
            this.label26.TabIndex = 32;
            this.label26.Text = "Author:";
            // 
            // edAuthor
            // 
            this.edAuthor.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edAuthor.Location = new System.Drawing.Point(56, 20);
            this.edAuthor.Name = "edAuthor";
            this.edAuthor.Size = new System.Drawing.Size(144, 20);
            this.edAuthor.TabIndex = 31;
            // 
            // label25
            // 
            this.label25.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label25.Location = new System.Drawing.Point(8, 0);
            this.label25.Name = "label25";
            this.label25.Size = new System.Drawing.Size(192, 16);
            this.label25.TabIndex = 20;
            this.label25.Text = "Pdf Properties:";
            // 
            // edFileName
            // 
            this.edFileName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.edFileName.BackColor = System.Drawing.Color.White;
            this.edFileName.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.edFileName.Location = new System.Drawing.Point(136, 16);
            this.edFileName.Name = "edFileName";
            this.edFileName.ReadOnly = true;
            this.edFileName.Size = new System.Drawing.Size(584, 13);
            this.edFileName.TabIndex = 4;
            this.edFileName.Text = "No file selected";
            // 
            // panel8
            // 
            this.panel8.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.panel8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel8.Controls.Add(this.chLandscape);
            this.panel8.Controls.Add(this.label24);
            this.panel8.Controls.Add(this.chPrintLeft);
            this.panel8.Controls.Add(this.chFormulaText);
            this.panel8.Controls.Add(this.chGridLines);
            this.panel8.Location = new System.Drawing.Point(32, 205);
            this.panel8.Name = "panel8";
            this.panel8.Size = new System.Drawing.Size(176, 120);
            this.panel8.TabIndex = 37;
            // 
            // chLandscape
            // 
            this.chLandscape.Location = new System.Drawing.Point(24, 88);
            this.chLandscape.Name = "chLandscape";
            this.chLandscape.Size = new System.Drawing.Size(136, 24);
            this.chLandscape.TabIndex = 20;
            this.chLandscape.Text = "Landscape";
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
            // chPrintLeft
            // 
            this.chPrintLeft.Location = new System.Drawing.Point(24, 47);
            this.chPrintLeft.Name = "chPrintLeft";
            this.chPrintLeft.Size = new System.Drawing.Size(152, 16);
            this.chPrintLeft.TabIndex = 18;
            this.chPrintLeft.Text = "Print Left, then down.";
            // 
            // chFormulaText
            // 
            this.chFormulaText.Location = new System.Drawing.Point(24, 71);
            this.chFormulaText.Name = "chFormulaText";
            this.chFormulaText.Size = new System.Drawing.Size(136, 16);
            this.chFormulaText.TabIndex = 17;
            this.chFormulaText.Text = "Print Formula Text";
            // 
            // chGridLines
            // 
            this.chGridLines.Location = new System.Drawing.Point(24, 24);
            this.chGridLines.Name = "chGridLines";
            this.chGridLines.Size = new System.Drawing.Size(128, 16);
            this.chGridLines.TabIndex = 16;
            this.chGridLines.Text = "Print Grid Lines";
            // 
            // panel7
            // 
            this.panel7.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel7.Controls.Add(this.edFooter);
            this.panel7.Controls.Add(this.label3);
            this.panel7.Controls.Add(this.edHeader);
            this.panel7.Controls.Add(this.label2);
            this.panel7.Controls.Add(this.label23);
            this.panel7.Location = new System.Drawing.Point(224, 333);
            this.panel7.Name = "panel7";
            this.panel7.Size = new System.Drawing.Size(296, 112);
            this.panel7.TabIndex = 36;
            // 
            // edFooter
            // 
            this.edFooter.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.edFooter.BackColor = System.Drawing.Color.White;
            this.edFooter.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edFooter.Location = new System.Drawing.Point(8, 88);
            this.edFooter.Name = "edFooter";
            this.edFooter.Size = new System.Drawing.Size(278, 20);
            this.edFooter.TabIndex = 46;
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.Color.White;
            this.label3.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(8, 72);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(56, 16);
            this.label3.TabIndex = 45;
            this.label3.Text = "Footer:";
            // 
            // edHeader
            // 
            this.edHeader.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.edHeader.BackColor = System.Drawing.Color.White;
            this.edHeader.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edHeader.Location = new System.Drawing.Point(8, 48);
            this.edHeader.Name = "edHeader";
            this.edHeader.Size = new System.Drawing.Size(278, 20);
            this.edHeader.TabIndex = 44;
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.Color.White;
            this.label2.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(8, 32);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(56, 16);
            this.label2.TabIndex = 43;
            this.label2.Text = "Header:";
            // 
            // label23
            // 
            this.label23.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label23.Location = new System.Drawing.Point(8, 8);
            this.label23.Name = "label23";
            this.label23.Size = new System.Drawing.Size(187, 16);
            this.label23.TabIndex = 42;
            this.label23.Text = "Headers and footers:";
            // 
            // panel6
            // 
            this.panel6.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(224)))), ((int)(((byte)(192)))));
            this.panel6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel6.Controls.Add(this.label22);
            this.panel6.Controls.Add(this.edf);
            this.panel6.Controls.Add(this.edb);
            this.panel6.Controls.Add(this.edr);
            this.panel6.Controls.Add(this.edt);
            this.panel6.Controls.Add(this.label7);
            this.panel6.Controls.Add(this.edl);
            this.panel6.Controls.Add(this.edh);
            this.panel6.Controls.Add(this.label9);
            this.panel6.Controls.Add(this.label10);
            this.panel6.Controls.Add(this.label8);
            this.panel6.Controls.Add(this.labelb);
            this.panel6.Controls.Add(this.label11);
            this.panel6.Location = new System.Drawing.Point(32, 333);
            this.panel6.Name = "panel6";
            this.panel6.Size = new System.Drawing.Size(176, 208);
            this.panel6.TabIndex = 35;
            // 
            // label22
            // 
            this.label22.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label22.Location = new System.Drawing.Point(8, 8);
            this.label22.Name = "label22";
            this.label22.Size = new System.Drawing.Size(120, 16);
            this.label22.TabIndex = 41;
            this.label22.Text = "Margins:";
            // 
            // edf
            // 
            this.edf.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edf.Location = new System.Drawing.Point(56, 152);
            this.edf.Name = "edf";
            this.edf.Size = new System.Drawing.Size(112, 20);
            this.edf.TabIndex = 39;
            // 
            // edb
            // 
            this.edb.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edb.Location = new System.Drawing.Point(56, 128);
            this.edb.Name = "edb";
            this.edb.Size = new System.Drawing.Size(112, 20);
            this.edb.TabIndex = 37;
            // 
            // edr
            // 
            this.edr.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edr.Location = new System.Drawing.Point(56, 56);
            this.edr.Name = "edr";
            this.edr.Size = new System.Drawing.Size(112, 20);
            this.edr.TabIndex = 35;
            // 
            // edt
            // 
            this.edt.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edt.Location = new System.Drawing.Point(56, 104);
            this.edt.Name = "edt";
            this.edt.Size = new System.Drawing.Size(112, 20);
            this.edt.TabIndex = 31;
            // 
            // label7
            // 
            this.label7.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(8, 32);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(36, 16);
            this.label7.TabIndex = 30;
            this.label7.Text = "Left:";
            // 
            // edl
            // 
            this.edl.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edl.Location = new System.Drawing.Point(56, 32);
            this.edl.Name = "edl";
            this.edl.Size = new System.Drawing.Size(112, 20);
            this.edl.TabIndex = 29;
            // 
            // edh
            // 
            this.edh.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edh.Location = new System.Drawing.Point(56, 80);
            this.edh.Name = "edh";
            this.edh.Size = new System.Drawing.Size(112, 20);
            this.edh.TabIndex = 32;
            // 
            // label9
            // 
            this.label9.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.Location = new System.Drawing.Point(8, 56);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(80, 16);
            this.label9.TabIndex = 36;
            this.label9.Text = "Right:";
            // 
            // label10
            // 
            this.label10.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.Location = new System.Drawing.Point(7, 80);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(88, 16);
            this.label10.TabIndex = 34;
            this.label10.Text = "Header:";
            // 
            // label8
            // 
            this.label8.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(8, 104);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(80, 16);
            this.label8.TabIndex = 33;
            this.label8.Text = "Top:";
            // 
            // labelb
            // 
            this.labelb.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelb.Location = new System.Drawing.Point(8, 130);
            this.labelb.Name = "labelb";
            this.labelb.Size = new System.Drawing.Size(88, 16);
            this.labelb.TabIndex = 38;
            this.labelb.Text = "Bottom:";
            // 
            // label11
            // 
            this.label11.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.Location = new System.Drawing.Point(8, 160);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(80, 16);
            this.label11.TabIndex = 40;
            this.label11.Text = "Footer:";
            // 
            // panel5
            // 
            this.panel5.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel5.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.panel5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel5.Controls.Add(this.label21);
            this.panel5.Controls.Add(this.label4);
            this.panel5.Controls.Add(this.edZoom);
            this.panel5.Controls.Add(this.chFitIn);
            this.panel5.Controls.Add(this.label6);
            this.panel5.Controls.Add(this.label5);
            this.panel5.Controls.Add(this.edVPages);
            this.panel5.Controls.Add(this.edHPages);
            this.panel5.Location = new System.Drawing.Point(224, 453);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(296, 88);
            this.panel5.TabIndex = 34;
            // 
            // label21
            // 
            this.label21.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label21.Location = new System.Drawing.Point(9, 5);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(192, 16);
            this.label21.TabIndex = 26;
            this.label21.Text = "Zoom:";
            // 
            // label4
            // 
            this.label4.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(72, 58);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(56, 16);
            this.label4.TabIndex = 25;
            this.label4.Text = "Zoom (%)";
            // 
            // edZoom
            // 
            this.edZoom.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edZoom.Location = new System.Drawing.Point(136, 56);
            this.edZoom.Name = "edZoom";
            this.edZoom.Size = new System.Drawing.Size(24, 20);
            this.edZoom.TabIndex = 24;
            // 
            // chFitIn
            // 
            this.chFitIn.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chFitIn.Location = new System.Drawing.Point(8, 24);
            this.chFitIn.Name = "chFitIn";
            this.chFitIn.Size = new System.Drawing.Size(56, 24);
            this.chFitIn.TabIndex = 23;
            this.chFitIn.Text = "Fit in";
            this.chFitIn.CheckedChanged += new System.EventHandler(this.chFitIn_CheckedChanged);
            // 
            // label6
            // 
            this.label6.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(208, 29);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(80, 16);
            this.label6.TabIndex = 22;
            this.label6.Text = "pages tall.";
            // 
            // label5
            // 
            this.label5.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(96, 29);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(80, 16);
            this.label5.TabIndex = 21;
            this.label5.Text = "pages wide x";
            // 
            // edVPages
            // 
            this.edVPages.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edVPages.Location = new System.Drawing.Point(176, 24);
            this.edVPages.Name = "edVPages";
            this.edVPages.ReadOnly = true;
            this.edVPages.Size = new System.Drawing.Size(24, 20);
            this.edVPages.TabIndex = 20;
            // 
            // edHPages
            // 
            this.edHPages.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.edHPages.Location = new System.Drawing.Point(64, 24);
            this.edHPages.Name = "edHPages";
            this.edHPages.ReadOnly = true;
            this.edHPages.Size = new System.Drawing.Size(24, 20);
            this.edHPages.TabIndex = 19;
            // 
            // panel4
            // 
            this.panel4.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.panel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel4.Controls.Add(this.chSubset);
            this.panel4.Controls.Add(this.cbKerning);
            this.panel4.Controls.Add(this.label20);
            this.panel4.Controls.Add(this.cbFontMapping);
            this.panel4.Controls.Add(this.chEmbed);
            this.panel4.Controls.Add(this.label19);
            this.panel4.Location = new System.Drawing.Point(224, 205);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(496, 120);
            this.panel4.TabIndex = 33;
            // 
            // chSubset
            // 
            this.chSubset.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.chSubset.Checked = true;
            this.chSubset.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chSubset.Location = new System.Drawing.Point(16, 72);
            this.chSubset.Name = "chSubset";
            this.chSubset.Size = new System.Drawing.Size(464, 16);
            this.chSubset.TabIndex = 36;
            this.chSubset.Text = "Subset fonts when embedding. (That is, embed only the characters used from the fo" +
    "nt)";
            // 
            // cbKerning
            // 
            this.cbKerning.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cbKerning.Location = new System.Drawing.Point(16, 88);
            this.cbKerning.Name = "cbKerning";
            this.cbKerning.Size = new System.Drawing.Size(464, 32);
            this.cbKerning.TabIndex = 35;
            this.cbKerning.Text = "Kerning. (Files with kerning look a little better but are a little bigger too)";
            // 
            // label20
            // 
            this.label20.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label20.Location = new System.Drawing.Point(16, 24);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(96, 16);
            this.label20.TabIndex = 34;
            this.label20.Text = "Font mapping:";
            // 
            // cbFontMapping
            // 
            this.cbFontMapping.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cbFontMapping.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbFontMapping.Items.AddRange(new object[] {
            "Replace all fonts by internal fonts. (smaller file size)",
            "Replace selected fonts by internal fonts. (optimum relation file size/accuracy)",
            "Do not replace any font. (maximum file size)"});
            this.cbFontMapping.Location = new System.Drawing.Point(120, 24);
            this.cbFontMapping.Name = "cbFontMapping";
            this.cbFontMapping.Size = new System.Drawing.Size(360, 21);
            this.cbFontMapping.TabIndex = 33;
            // 
            // chEmbed
            // 
            this.chEmbed.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.chEmbed.Checked = true;
            this.chEmbed.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chEmbed.Location = new System.Drawing.Point(16, 40);
            this.chEmbed.Name = "chEmbed";
            this.chEmbed.Size = new System.Drawing.Size(464, 32);
            this.chEmbed.TabIndex = 3;
            this.chEmbed.Text = "Embed all fonts. (if you leave this option off, some fonts might be embedded anyw" +
    "ay)";
            // 
            // label19
            // 
            this.label19.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label19.Location = new System.Drawing.Point(3, 8);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(192, 16);
            this.label19.TabIndex = 2;
            this.label19.Text = "Fonts:";
            // 
            // label18
            // 
            this.label18.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label18.Location = new System.Drawing.Point(168, 40);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(96, 16);
            this.label18.TabIndex = 32;
            this.label18.Text = "Sheet to export:";
            // 
            // cbSheet
            // 
            this.cbSheet.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cbSheet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbSheet.Location = new System.Drawing.Point(264, 34);
            this.cbSheet.Name = "cbSheet";
            this.cbSheet.Size = new System.Drawing.Size(256, 21);
            this.cbSheet.TabIndex = 31;
            this.cbSheet.SelectedIndexChanged += new System.EventHandler(this.cbSheet_SelectedIndexChanged);
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
            this.panel3.Location = new System.Drawing.Point(536, 333);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(184, 208);
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
            this.label13.Size = new System.Drawing.Size(160, 32);
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
            this.label12.Text = "Range to Export:";
            // 
            // chExportAll
            // 
            this.chExportAll.Location = new System.Drawing.Point(32, 40);
            this.chExportAll.Name = "chExportAll";
            this.chExportAll.Size = new System.Drawing.Size(128, 16);
            this.chExportAll.TabIndex = 1;
            this.chExportAll.Text = "Export all Sheets";
            this.chExportAll.CheckedChanged += new System.EventHandler(this.chExportAll_CheckedChanged);
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(29, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(80, 16);
            this.label1.TabIndex = 0;
            this.label1.Text = "File to print:";
            // 
            // exportDialog
            // 
            this.exportDialog.DefaultExt = "pdf";
            this.exportDialog.Filter = "Pdf files|*.pdf";
            // 
            // mainToolbar
            // 
            this.mainToolbar.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openFile,
            this.export,
            this.btnClose});
            this.mainToolbar.Location = new System.Drawing.Point(0, 0);
            this.mainToolbar.Name = "mainToolbar";
            this.mainToolbar.Size = new System.Drawing.Size(768, 38);
            this.mainToolbar.TabIndex = 4;
            this.mainToolbar.Text = "toolStrip1";
            // 
            // openFile
            // 
            this.openFile.Image = ((System.Drawing.Image)(resources.GetObject("openFile.Image")));
            this.openFile.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.openFile.Name = "openFile";
            this.openFile.Size = new System.Drawing.Size(61, 35);
            this.openFile.Text = "Open File";
            this.openFile.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.openFile.Click += new System.EventHandler(this.openFile_Click);
            // 
            // export
            // 
            this.export.Image = ((System.Drawing.Image)(resources.GetObject("export.Image")));
            this.export.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.export.Name = "export";
            this.export.Size = new System.Drawing.Size(82, 35);
            this.export.Text = "Export to PDF";
            this.export.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.export.Click += new System.EventHandler(this.export_Click);
            // 
            // btnClose
            // 
            this.btnClose.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.btnClose.Image = ((System.Drawing.Image)(resources.GetObject("btnClose.Image")));
            this.btnClose.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(59, 35);
            this.btnClose.Text = "     E&xit     ";
            this.btnClose.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnClose.Click += new System.EventHandler(this.button2_Click);
            // 
            // flexCelPdfExport1
            // 
            this.flexCelPdfExport1.FontEmbed = FlexCel.Pdf.TFontEmbed.Embed;
            this.flexCelPdfExport1.InitialZoomAndPage = null;
            this.flexCelPdfExport1.PageLayout = FlexCel.Pdf.TPageLayout.None;
            this.flexCelPdfExport1.PageLayoutDisplay = FlexCel.Pdf.TPageLayoutDisplay.None;
            this.flexCelPdfExport1.PageSize = null;
            tPdfProperties1.Author = null;
            tPdfProperties1.Creator = null;
            tPdfProperties1.Keywords = null;
            tPdfProperties1.Language = null;
            tPdfProperties1.Subject = null;
            tPdfProperties1.Title = null;
            this.flexCelPdfExport1.Properties = tPdfProperties1;
            this.flexCelPdfExport1.TagMode = FlexCel.Pdf.TTagMode.Full;
            this.flexCelPdfExport1.UnlicensedReplacementFont = null;
            this.flexCelPdfExport1.UseExcelProperties = true;
            this.flexCelPdfExport1.Workbook = null;
            this.flexCelPdfExport1.AfterGeneratePage += new FlexCel.Render.PageEventHandler(this.flexCelPdfExport1_AfterGeneratePage);
            this.flexCelPdfExport1.GetFontData += new FlexCel.Pdf.GetFontDataEventHandler(this.flexCelPdfExport1_GetFontData);
            // 
            // mainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(768, 631);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.mainToolbar);
            this.Name = "mainForm";
            this.Text = "Export an Excel file to pdf";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel9.ResumeLayout(false);
            this.panel9.PerformLayout();
            this.panel8.ResumeLayout(false);
            this.panel7.ResumeLayout(false);
            this.panel7.PerformLayout();
            this.panel6.ResumeLayout(false);
            this.panel6.PerformLayout();
            this.panel5.ResumeLayout(false);
            this.panel5.PerformLayout();
            this.panel4.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
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
        private Label label29;
        private TextBox edLang;
        private Panel panel2;
        private ComboBox cbPdfType;
        private Label label34;
        private ComboBox cbVersion;
        private Label label31;
        private Label label30;
        private ComboBox cbTagged;
        private Label label32;
    }
}

