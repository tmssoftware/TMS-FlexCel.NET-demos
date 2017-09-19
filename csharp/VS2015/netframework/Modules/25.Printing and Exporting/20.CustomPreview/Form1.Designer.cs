using System;
using System.Drawing;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Drawing.Drawing2D;
using System.Diagnostics;
using System.Threading;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using FlexCel.Winforms;
using FlexCel.Render;
using FlexCel.Pdf;
namespace CustomPreview
{
    public partial class mainForm: System.Windows.Forms.Form
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(mainForm));
            this.flexCelImgExport1 = new FlexCel.Render.FlexCelImgExport();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.panel1 = new System.Windows.Forms.Panel();
            this.MainPreview = new FlexCel.Winforms.FlexCelPreview();
            this.thumbs = new FlexCel.Winforms.FlexCelPreview();
            this.splitter1 = new System.Windows.Forms.Splitter();
            this.panelLeft = new System.Windows.Forms.Panel();
            this.cbAllSheets = new System.Windows.Forms.CheckBox();
            this.label2 = new System.Windows.Forms.Label();
            this.sheetSplitter = new System.Windows.Forms.Splitter();
            this.lbSheets = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.PdfSaveFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.mainToolbar = new System.Windows.Forms.ToolStrip();
            this.openFile = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.btnFirst = new System.Windows.Forms.ToolStripButton();
            this.btnPrev = new System.Windows.Forms.ToolStripButton();
            this.edPage = new System.Windows.Forms.ToolStripTextBox();
            this.btnNext = new System.Windows.Forms.ToolStripButton();
            this.btnLast = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.btnAutofit = new System.Windows.Forms.ToolStripDropDownButton();
            this.noneToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.fitToWidthToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.fitToHeightToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.fitToPageToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.btnZoomOut = new System.Windows.Forms.ToolStripButton();
            this.edZoom = new System.Windows.Forms.ToolStripTextBox();
            this.btnZoomIn = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
            this.btnGridLines = new System.Windows.Forms.ToolStripButton();
            this.btnHeadings = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator4 = new System.Windows.Forms.ToolStripSeparator();
            this.btnRecalc = new System.Windows.Forms.ToolStripButton();
            this.btnPdf = new System.Windows.Forms.ToolStripButton();
            this.btnClose = new System.Windows.Forms.ToolStripButton();
            this.panel1.SuspendLayout();
            this.panelLeft.SuspendLayout();
            this.mainToolbar.SuspendLayout();
            this.SuspendLayout();
            // 
            // flexCelImgExport1
            // 
            this.flexCelImgExport1.AllVisibleSheets = false;
            this.flexCelImgExport1.PageSize = null;
            this.flexCelImgExport1.ResetPageNumberOnEachSheet = false;
            this.flexCelImgExport1.Resolution = 96D;
            this.flexCelImgExport1.Workbook = null;
            // 
            // openFileDialog
            // 
            this.openFileDialog.DefaultExt = "xls";
            this.openFileDialog.Filter = "Excel Files|*.xls|All files|*.*";
            this.openFileDialog.Title = "Select a file to preview";
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.MainPreview);
            this.panel1.Controls.Add(this.splitter1);
            this.panel1.Controls.Add(this.panelLeft);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 46);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(808, 375);
            this.panel1.TabIndex = 8;
            // 
            // MainPreview
            // 
            this.MainPreview.AutoScrollMinSize = new System.Drawing.Size(40, 383);
            this.MainPreview.Dock = System.Windows.Forms.DockStyle.Fill;
            this.MainPreview.Document = this.flexCelImgExport1;
            this.MainPreview.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            this.MainPreview.Location = new System.Drawing.Point(144, 0);
            this.MainPreview.Name = "MainPreview";
            this.MainPreview.PageXSeparation = 20;
            this.MainPreview.Size = new System.Drawing.Size(662, 373);
            this.MainPreview.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            this.MainPreview.StartPage = 1;
            this.MainPreview.TabIndex = 2;
            this.MainPreview.ThumbnailLarge = null;
            this.MainPreview.ThumbnailSmall = this.thumbs;
            this.MainPreview.StartPageChanged += new System.EventHandler(this.flexCelPreview1_StartPageChanged);
            this.MainPreview.ZoomChanged += new System.EventHandler(this.flexCelPreview1_ZoomChanged);
            // 
            // thumbs
            // 
            this.thumbs.AutoScrollMinSize = new System.Drawing.Size(20, 10);
            this.thumbs.Dock = System.Windows.Forms.DockStyle.Fill;
            this.thumbs.Document = this.flexCelImgExport1;
            this.thumbs.Location = new System.Drawing.Point(0, 115);
            this.thumbs.Name = "thumbs";
            this.thumbs.Size = new System.Drawing.Size(136, 258);
            this.thumbs.StartPage = 1;
            this.thumbs.TabIndex = 3;
            this.thumbs.ThumbnailLarge = this.MainPreview;
            this.thumbs.ThumbnailSmall = null;
            this.thumbs.Zoom = 0.1D;
            // 
            // splitter1
            // 
            this.splitter1.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.splitter1.Location = new System.Drawing.Point(136, 0);
            this.splitter1.MinSize = 0;
            this.splitter1.Name = "splitter1";
            this.splitter1.Size = new System.Drawing.Size(8, 373);
            this.splitter1.TabIndex = 11;
            this.splitter1.TabStop = false;
            // 
            // panelLeft
            // 
            this.panelLeft.Controls.Add(this.cbAllSheets);
            this.panelLeft.Controls.Add(this.thumbs);
            this.panelLeft.Controls.Add(this.label2);
            this.panelLeft.Controls.Add(this.sheetSplitter);
            this.panelLeft.Controls.Add(this.lbSheets);
            this.panelLeft.Controls.Add(this.label1);
            this.panelLeft.Dock = System.Windows.Forms.DockStyle.Left;
            this.panelLeft.Location = new System.Drawing.Point(0, 0);
            this.panelLeft.Name = "panelLeft";
            this.panelLeft.Size = new System.Drawing.Size(136, 373);
            this.panelLeft.TabIndex = 9;
            // 
            // cbAllSheets
            // 
            this.cbAllSheets.Location = new System.Drawing.Point(16, 16);
            this.cbAllSheets.Name = "cbAllSheets";
            this.cbAllSheets.Size = new System.Drawing.Size(104, 16);
            this.cbAllSheets.TabIndex = 14;
            this.cbAllSheets.Text = "All Sheets";
            this.cbAllSheets.CheckedChanged += new System.EventHandler(this.cbAllSheets_CheckedChanged);
            // 
            // label2
            // 
            this.label2.Dock = System.Windows.Forms.DockStyle.Top;
            this.label2.Location = new System.Drawing.Point(0, 99);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(136, 16);
            this.label2.TabIndex = 13;
            this.label2.Text = "Thumbs";
            // 
            // sheetSplitter
            // 
            this.sheetSplitter.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.sheetSplitter.Dock = System.Windows.Forms.DockStyle.Top;
            this.sheetSplitter.Location = new System.Drawing.Point(0, 91);
            this.sheetSplitter.Name = "sheetSplitter";
            this.sheetSplitter.Size = new System.Drawing.Size(136, 8);
            this.sheetSplitter.TabIndex = 11;
            this.sheetSplitter.TabStop = false;
            // 
            // lbSheets
            // 
            this.lbSheets.Dock = System.Windows.Forms.DockStyle.Top;
            this.lbSheets.Items.AddRange(new object[] {
            "No open file"});
            this.lbSheets.Location = new System.Drawing.Point(0, 35);
            this.lbSheets.Name = "lbSheets";
            this.lbSheets.Size = new System.Drawing.Size(136, 56);
            this.lbSheets.TabIndex = 10;
            this.lbSheets.SelectedIndexChanged += new System.EventHandler(this.lbSheets_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.Dock = System.Windows.Forms.DockStyle.Top;
            this.label1.Location = new System.Drawing.Point(0, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(136, 35);
            this.label1.TabIndex = 12;
            this.label1.Text = "Sheets";
            // 
            // PdfSaveFileDialog
            // 
            this.PdfSaveFileDialog.DefaultExt = "pdf";
            this.PdfSaveFileDialog.Filter = "Pdf Files|*.pdf";
            this.PdfSaveFileDialog.Title = "Select the file to export to:";
            // 
            // mainToolbar
            // 
            this.mainToolbar.ImageScalingSize = new System.Drawing.Size(24, 24);
            this.mainToolbar.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openFile,
            this.toolStripSeparator2,
            this.btnFirst,
            this.btnPrev,
            this.edPage,
            this.btnNext,
            this.btnLast,
            this.toolStripSeparator1,
            this.btnAutofit,
            this.btnZoomOut,
            this.edZoom,
            this.btnZoomIn,
            this.toolStripSeparator3,
            this.btnGridLines,
            this.btnHeadings,
            this.toolStripSeparator4,
            this.btnRecalc,
            this.btnPdf,
            this.btnClose});
            this.mainToolbar.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.HorizontalStackWithOverflow;
            this.mainToolbar.Location = new System.Drawing.Point(0, 0);
            this.mainToolbar.Name = "mainToolbar";
            this.mainToolbar.Size = new System.Drawing.Size(808, 46);
            this.mainToolbar.TabIndex = 14;
            // 
            // openFile
            // 
            this.openFile.Image = global::CustomPreview.Properties.Resources.open;
            this.openFile.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.openFile.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.openFile.Name = "openFile";
            this.openFile.Size = new System.Drawing.Size(61, 43);
            this.openFile.Text = "&Open File";
            this.openFile.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.openFile.ToolTipText = "Open an Excel file";
            this.openFile.Click += new System.EventHandler(this.openFile_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.AutoSize = false;
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(20, 46);
            // 
            // btnFirst
            // 
            this.btnFirst.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.btnFirst.Enabled = false;
            this.btnFirst.Image = ((System.Drawing.Image)(resources.GetObject("btnFirst.Image")));
            this.btnFirst.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnFirst.Name = "btnFirst";
            this.btnFirst.Size = new System.Drawing.Size(27, 43);
            this.btnFirst.Text = "<<";
            this.btnFirst.ToolTipText = "First page";
            this.btnFirst.Click += new System.EventHandler(this.btnFirst_Click);
            // 
            // btnPrev
            // 
            this.btnPrev.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.btnPrev.Enabled = false;
            this.btnPrev.Image = ((System.Drawing.Image)(resources.GetObject("btnPrev.Image")));
            this.btnPrev.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnPrev.Name = "btnPrev";
            this.btnPrev.Size = new System.Drawing.Size(23, 43);
            this.btnPrev.Text = "<";
            this.btnPrev.ToolTipText = "Previous page";
            this.btnPrev.Click += new System.EventHandler(this.btnPrev_Click);
            // 
            // edPage
            // 
            this.edPage.AutoSize = false;
            this.edPage.Enabled = false;
            this.edPage.Name = "edPage";
            this.edPage.Size = new System.Drawing.Size(100, 18);
            this.edPage.TextBoxTextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.edPage.Leave += new System.EventHandler(this.edPage_Leave);
            this.edPage.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.edPage_KeyPress);
            // 
            // btnNext
            // 
            this.btnNext.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.btnNext.Enabled = false;
            this.btnNext.Image = ((System.Drawing.Image)(resources.GetObject("btnNext.Image")));
            this.btnNext.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnNext.Name = "btnNext";
            this.btnNext.Size = new System.Drawing.Size(23, 43);
            this.btnNext.Text = ">";
            this.btnNext.ToolTipText = "Next page";
            this.btnNext.Click += new System.EventHandler(this.btnNext_Click);
            // 
            // btnLast
            // 
            this.btnLast.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.btnLast.Enabled = false;
            this.btnLast.Image = ((System.Drawing.Image)(resources.GetObject("btnLast.Image")));
            this.btnLast.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnLast.Name = "btnLast";
            this.btnLast.Size = new System.Drawing.Size(27, 43);
            this.btnLast.Text = ">>";
            this.btnLast.ToolTipText = "Last page";
            this.btnLast.Click += new System.EventHandler(this.btnLast_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 46);
            // 
            // btnAutofit
            // 
            this.btnAutofit.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.noneToolStripMenuItem,
            this.fitToWidthToolStripMenuItem,
            this.fitToHeightToolStripMenuItem,
            this.fitToPageToolStripMenuItem});
            this.btnAutofit.Image = global::CustomPreview.Properties.Resources.autofit;
            this.btnAutofit.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnAutofit.Name = "btnAutofit";
            this.btnAutofit.Size = new System.Drawing.Size(76, 43);
            this.btnAutofit.Text = "No Autofit";
            this.btnAutofit.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            // 
            // noneToolStripMenuItem
            // 
            this.noneToolStripMenuItem.Name = "noneToolStripMenuItem";
            this.noneToolStripMenuItem.Size = new System.Drawing.Size(140, 22);
            this.noneToolStripMenuItem.Text = "No Autofit";
            this.noneToolStripMenuItem.Click += new System.EventHandler(this.noneToolStripMenuItem_Click);
            // 
            // fitToWidthToolStripMenuItem
            // 
            this.fitToWidthToolStripMenuItem.Name = "fitToWidthToolStripMenuItem";
            this.fitToWidthToolStripMenuItem.Size = new System.Drawing.Size(140, 22);
            this.fitToWidthToolStripMenuItem.Text = "Fit to Width";
            this.fitToWidthToolStripMenuItem.Click += new System.EventHandler(this.fitToWidthToolStripMenuItem_Click);
            // 
            // fitToHeightToolStripMenuItem
            // 
            this.fitToHeightToolStripMenuItem.Name = "fitToHeightToolStripMenuItem";
            this.fitToHeightToolStripMenuItem.Size = new System.Drawing.Size(140, 22);
            this.fitToHeightToolStripMenuItem.Text = "Fit to Height";
            this.fitToHeightToolStripMenuItem.Click += new System.EventHandler(this.fitToHeightToolStripMenuItem_Click);
            // 
            // fitToPageToolStripMenuItem
            // 
            this.fitToPageToolStripMenuItem.Name = "fitToPageToolStripMenuItem";
            this.fitToPageToolStripMenuItem.Size = new System.Drawing.Size(140, 22);
            this.fitToPageToolStripMenuItem.Text = "Fit to Page";
            this.fitToPageToolStripMenuItem.Click += new System.EventHandler(this.fitToPageToolStripMenuItem_Click);
            // 
            // btnZoomOut
            // 
            this.btnZoomOut.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.btnZoomOut.Enabled = false;
            this.btnZoomOut.Image = ((System.Drawing.Image)(resources.GetObject("btnZoomOut.Image")));
            this.btnZoomOut.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnZoomOut.Name = "btnZoomOut";
            this.btnZoomOut.Size = new System.Drawing.Size(23, 43);
            this.btnZoomOut.Text = "-";
            this.btnZoomOut.ToolTipText = "Zoom out";
            this.btnZoomOut.Click += new System.EventHandler(this.btnZoomOut_Click);
            // 
            // edZoom
            // 
            this.edZoom.AutoSize = false;
            this.edZoom.Enabled = false;
            this.edZoom.Name = "edZoom";
            this.edZoom.Size = new System.Drawing.Size(40, 18);
            this.edZoom.TextBoxTextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.edZoom.Enter += new System.EventHandler(this.edZoom_Enter);
            this.edZoom.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.edZoom_KeyPress);
            // 
            // btnZoomIn
            // 
            this.btnZoomIn.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.btnZoomIn.Enabled = false;
            this.btnZoomIn.Image = ((System.Drawing.Image)(resources.GetObject("btnZoomIn.Image")));
            this.btnZoomIn.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnZoomIn.Name = "btnZoomIn";
            this.btnZoomIn.Size = new System.Drawing.Size(23, 43);
            this.btnZoomIn.Text = "+";
            this.btnZoomIn.ToolTipText = "Zoom in";
            this.btnZoomIn.Click += new System.EventHandler(this.btnZoomIn_Click);
            // 
            // toolStripSeparator3
            // 
            this.toolStripSeparator3.AutoSize = false;
            this.toolStripSeparator3.Name = "toolStripSeparator3";
            this.toolStripSeparator3.Size = new System.Drawing.Size(20, 46);
            // 
            // btnGridLines
            // 
            this.btnGridLines.CheckOnClick = true;
            this.btnGridLines.Enabled = false;
            this.btnGridLines.Image = global::CustomPreview.Properties.Resources.grid;
            this.btnGridLines.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnGridLines.Name = "btnGridLines";
            this.btnGridLines.Size = new System.Drawing.Size(57, 43);
            this.btnGridLines.Text = "&Gridlines";
            this.btnGridLines.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnGridLines.ToolTipText = "Show gridlines";
            this.btnGridLines.Click += new System.EventHandler(this.btnGridLines_Click);
            // 
            // btnHeadings
            // 
            this.btnHeadings.CheckOnClick = true;
            this.btnHeadings.Enabled = false;
            this.btnHeadings.Image = global::CustomPreview.Properties.Resources.Head;
            this.btnHeadings.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnHeadings.Name = "btnHeadings";
            this.btnHeadings.Size = new System.Drawing.Size(61, 43);
            this.btnHeadings.Text = "&Headings";
            this.btnHeadings.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnHeadings.ToolTipText = "Show the headings";
            this.btnHeadings.Click += new System.EventHandler(this.btnHeadings_Click);
            // 
            // toolStripSeparator4
            // 
            this.toolStripSeparator4.Name = "toolStripSeparator4";
            this.toolStripSeparator4.Size = new System.Drawing.Size(6, 46);
            // 
            // btnRecalc
            // 
            this.btnRecalc.Enabled = false;
            this.btnRecalc.Image = global::CustomPreview.Properties.Resources.calc;
            this.btnRecalc.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.btnRecalc.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnRecalc.Name = "btnRecalc";
            this.btnRecalc.Size = new System.Drawing.Size(45, 43);
            this.btnRecalc.Text = "&Recalc";
            this.btnRecalc.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnRecalc.ToolTipText = "Recalculate the file";
            this.btnRecalc.Click += new System.EventHandler(this.btnRecalc_Click);
            // 
            // btnPdf
            // 
            this.btnPdf.Enabled = false;
            this.btnPdf.Image = global::CustomPreview.Properties.Resources.pdf;
            this.btnPdf.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.btnPdf.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnPdf.Name = "btnPdf";
            this.btnPdf.Size = new System.Drawing.Size(79, 43);
            this.btnPdf.Text = "Export to &Pdf";
            this.btnPdf.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnPdf.ToolTipText = "Export the file to Pdf";
            this.btnPdf.Click += new System.EventHandler(this.btnPdf_Click);
            // 
            // btnClose
            // 
            this.btnClose.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.btnClose.Image = global::CustomPreview.Properties.Resources.close;
            this.btnClose.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.btnClose.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(59, 43);
            this.btnClose.Text = "     E&xit     ";
            this.btnClose.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnClose.ToolTipText = "Exit from the application";
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // mainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(808, 421);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.mainToolbar);
            this.Name = "mainForm";
            this.Text = "Custom Preview Demo";
            this.Load += new System.EventHandler(this.mainForm_Load);
            this.panel1.ResumeLayout(false);
            this.panelLeft.ResumeLayout(false);
            this.mainToolbar.ResumeLayout(false);
            this.mainToolbar.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panelLeft;
        private System.Windows.Forms.ListBox lbSheets;
        private System.Windows.Forms.Splitter splitter1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.SaveFileDialog PdfSaveFileDialog;
        private System.Windows.Forms.CheckBox cbAllSheets;
        private System.Windows.Forms.Splitter sheetSplitter;
        private System.Windows.Forms.ToolTip toolTip1;
        private FlexCel.Render.FlexCelImgExport flexCelImgExport1;
        private FlexCel.Winforms.FlexCelPreview MainPreview;
        private FlexCel.Winforms.FlexCelPreview thumbs;
        private ToolStrip mainToolbar;
        private ToolStripButton openFile;
        private ToolStripSeparator toolStripSeparator1;
        private ToolStripButton btnRecalc;
        private ToolStripButton btnPdf;
        private ToolStripButton btnClose;
        private ToolStripSeparator toolStripSeparator2;
        private ToolStripButton btnFirst;
        private ToolStripButton btnPrev;
        private ToolStripTextBox edPage;
        private ToolStripButton btnNext;
        private ToolStripButton btnLast;
        private ToolStripButton btnZoomOut;
        private ToolStripTextBox edZoom;
        private ToolStripButton btnZoomIn;
        private ToolStripSeparator toolStripSeparator3;
        private ToolStripButton btnHeadings;
        private ToolStripButton btnGridLines;
        private ToolStripSeparator toolStripSeparator4;
        private ToolStripDropDownButton btnAutofit;
        private ToolStripMenuItem noneToolStripMenuItem;
        private ToolStripMenuItem fitToWidthToolStripMenuItem;
        private ToolStripMenuItem fitToHeightToolStripMenuItem;
        private ToolStripMenuItem fitToPageToolStripMenuItem;
    }
}

