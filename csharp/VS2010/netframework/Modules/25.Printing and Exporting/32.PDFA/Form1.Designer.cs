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
namespace PDFA
{
    public partial class mainForm: System.Windows.Forms.Form
    {
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.SaveFileDialog exportDialog;
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.cbEmbed = new System.Windows.Forms.CheckBox();
            this.cbPdfType = new System.Windows.Forms.ComboBox();
            this.exportDialog = new System.Windows.Forms.SaveFileDialog();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.export = new System.Windows.Forms.ToolStripButton();
            this.btnClose = new System.Windows.Forms.ToolStripButton();
            this.panel1.SuspendLayout();
            this.toolStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.Controls.Add(this.cbEmbed);
            this.panel1.Controls.Add(this.cbPdfType);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 38);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(372, 104);
            this.panel1.TabIndex = 3;
            // 
            // cbEmbed
            // 
            this.cbEmbed.AutoSize = true;
            this.cbEmbed.Location = new System.Drawing.Point(12, 54);
            this.cbEmbed.Name = "cbEmbed";
            this.cbEmbed.Size = new System.Drawing.Size(351, 17);
            this.cbEmbed.TabIndex = 5;
            this.cbEmbed.Text = "Embed xslx source file inside the PDF. (requires PDF/A3 or Standard)";
            this.cbEmbed.UseVisualStyleBackColor = true;
            // 
            // cbPdfType
            // 
            this.cbPdfType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbPdfType.Items.AddRange(new object[] {
            "Standard",
            "PDF/A1a",
            "PDF/A1b",
            "PDF/A2a",
            "PDF/A2b",
            "PDF/A3a",
            "PDF/A3b"});
            this.cbPdfType.Location = new System.Drawing.Point(12, 14);
            this.cbPdfType.Name = "cbPdfType";
            this.cbPdfType.Size = new System.Drawing.Size(144, 21);
            this.cbPdfType.TabIndex = 35;
            // 
            // exportDialog
            // 
            this.exportDialog.DefaultExt = "pdf";
            this.exportDialog.Filter = "Pdf files|*.pdf";
            // 
            // toolStrip1
            // 
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.export,
            this.btnClose});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(372, 38);
            this.toolStrip1.TabIndex = 4;
            this.toolStrip1.Text = "mainToolbar";
            // 
            // export
            // 
            this.export.Image = ((System.Drawing.Image)(resources.GetObject("export.Image")));
            this.export.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.export.Name = "export";
            this.export.Size = new System.Drawing.Size(69, 35);
            this.export.Text = "Create PDF";
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
            // mainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(372, 142);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.toolStrip1);
            this.Name = "mainForm";
            this.Text = "PDF/A";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        private ToolStrip toolStrip1;
        private ToolStripButton export;
        private ToolStripButton btnClose;
        private CheckBox cbEmbed;
        private ComboBox cbPdfType;
    }
}

