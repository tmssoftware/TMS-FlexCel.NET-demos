using System;
using System.Drawing;
using System.Drawing.Drawing2D;
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
using System.Text;
namespace RenderObjects
{
    public partial class mainForm: System.Windows.Forms.Form
    {
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.panelError = new System.Windows.Forms.Panel();
            this.labelError = new System.Windows.Forms.Label();
            this.chartBox = new System.Windows.Forms.PictureBox();
            this.panel7 = new System.Windows.Forms.Panel();
            this.cbTheme = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.checkBox4 = new System.Windows.Forms.CheckBox();
            this.updater = new System.Windows.Forms.Timer(this.components);
            this.mainToolbar = new System.Windows.Forms.ToolStrip();
            this.btnRun = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.btnExit = new System.Windows.Forms.ToolStripButton();
            this.btnCancel = new System.Windows.Forms.ToolStripButton();
            this.panel1.SuspendLayout();
            this.panelError.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.chartBox)).BeginInit();
            this.panel7.SuspendLayout();
            this.mainToolbar.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.Controls.Add(this.panelError);
            this.panel1.Controls.Add(this.chartBox);
            this.panel1.Controls.Add(this.panel7);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 38);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(464, 392);
            this.panel1.TabIndex = 3;
            // 
            // panelError
            // 
            this.panelError.Controls.Add(this.labelError);
            this.panelError.Location = new System.Drawing.Point(136, 128);
            this.panelError.Name = "panelError";
            this.panelError.Size = new System.Drawing.Size(200, 100);
            this.panelError.TabIndex = 52;
            this.panelError.Visible = false;
            // 
            // labelError
            // 
            this.labelError.Location = new System.Drawing.Point(8, 16);
            this.labelError.Name = "labelError";
            this.labelError.Size = new System.Drawing.Size(100, 23);
            this.labelError.TabIndex = 0;
            // 
            // chartBox
            // 
            this.chartBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.chartBox.Location = new System.Drawing.Point(24, 120);
            this.chartBox.Name = "chartBox";
            this.chartBox.Size = new System.Drawing.Size(416, 250);
            this.chartBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.chartBox.TabIndex = 51;
            this.chartBox.TabStop = false;
            // 
            // panel7
            // 
            this.panel7.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel7.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.panel7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel7.Controls.Add(this.cbTheme);
            this.panel7.Controls.Add(this.label2);
            this.panel7.Location = new System.Drawing.Point(16, 16);
            this.panel7.Name = "panel7";
            this.panel7.Size = new System.Drawing.Size(432, 72);
            this.panel7.TabIndex = 44;
            // 
            // cbTheme
            // 
            this.cbTheme.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbTheme.Location = new System.Drawing.Point(8, 32);
            this.cbTheme.Name = "cbTheme";
            this.cbTheme.Size = new System.Drawing.Size(248, 21);
            this.cbTheme.TabIndex = 46;
            this.cbTheme.SelectedIndexChanged += new System.EventHandler(this.cbTheme_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(8, 8);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(192, 16);
            this.label2.TabIndex = 19;
            this.label2.Text = "Select Theme:";
            // 
            // checkBox4
            // 
            this.checkBox4.Location = new System.Drawing.Point(0, 0);
            this.checkBox4.Name = "checkBox4";
            this.checkBox4.Size = new System.Drawing.Size(104, 24);
            this.checkBox4.TabIndex = 0;
            // 
            // updater
            // 
            this.updater.Tick += new System.EventHandler(this.updater_Tick);
            // 
            // mainToolbar
            // 
            this.mainToolbar.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnRun,
            this.toolStripSeparator1,
            this.btnExit,
            this.btnCancel});
            this.mainToolbar.Location = new System.Drawing.Point(0, 0);
            this.mainToolbar.Name = "mainToolbar";
            this.mainToolbar.Size = new System.Drawing.Size(464, 38);
            this.mainToolbar.TabIndex = 11;
            this.mainToolbar.Text = "toolStrip1";
            // 
            // btnRun
            // 
            this.btnRun.Image = ((System.Drawing.Image)(resources.GetObject("btnRun.Image")));
            this.btnRun.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(35, 35);
            this.btnRun.Text = "Run!";
            this.btnRun.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 38);
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
            // btnCancel
            // 
            this.btnCancel.Enabled = false;
            this.btnCancel.Image = ((System.Drawing.Image)(resources.GetObject("btnCancel.Image")));
            this.btnCancel.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(47, 35);
            this.btnCancel.Text = "Cancel";
            this.btnCancel.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // mainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(464, 430);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.mainToolbar);
            this.MaximumSize = new System.Drawing.Size(800, 800);
            this.Name = "mainForm";
            this.Text = "Using FlexCel to render just a part of a spreadshet";
            this.panel1.ResumeLayout(false);
            this.panelError.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.chartBox)).EndInit();
            this.panel7.ResumeLayout(false);
            this.mainToolbar.ResumeLayout(false);
            this.mainToolbar.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.CheckBox checkBox4;
        private System.Windows.Forms.Panel panel7;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.PictureBox chartBox;
        private System.Windows.Forms.Timer updater;
        private System.Windows.Forms.ComboBox cbTheme;
        private System.Windows.Forms.Panel panelError;
        private System.Windows.Forms.Label labelError;
        private ToolStrip mainToolbar;
        private ToolStripButton btnRun;
        private ToolStripSeparator toolStripSeparator1;
        private ToolStripButton btnExit;
        private ToolStripButton btnCancel;
    }
}

