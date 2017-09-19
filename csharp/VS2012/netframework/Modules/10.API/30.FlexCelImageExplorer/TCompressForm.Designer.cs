using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.IO;
namespace FlexCelImageExplorer
{
    public partial class TCompressForm: System.Windows.Forms.Form
    {
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.NumericUpDown edPercent;
        private System.Windows.Forms.ComboBox cbPixelFormat;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.CheckBox cbTransparent;
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.Container components = null;

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
            this.label1 = new System.Windows.Forms.Label();
            this.edPercent = new System.Windows.Forms.NumericUpDown();
            this.cbPixelFormat = new System.Windows.Forms.ComboBox();
            this.btnOk = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.cbTransparent = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.edPercent)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(16, 8);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(144, 16);
            this.label1.TabIndex = 0;
            this.label1.Text = "Change Resolution (%):";
            // 
            // edPercent
            // 
            this.edPercent.Increment = new System.Decimal(new int[] {
                                                                        5,
                                                                        0,
                                                                        0,
                                                                        0});
            this.edPercent.Location = new System.Drawing.Point(176, 8);
            this.edPercent.Minimum = new System.Decimal(new int[] {
                                                                      10,
                                                                      0,
                                                                      0,
                                                                      0});
            this.edPercent.Name = "edPercent";
            this.edPercent.Size = new System.Drawing.Size(48, 20);
            this.edPercent.TabIndex = 2;
            this.edPercent.Value = new System.Decimal(new int[] {
                                                                    60,
                                                                    0,
                                                                    0,
                                                                    0});
            // 
            // cbPixelFormat
            // 
            this.cbPixelFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbPixelFormat.Items.AddRange(new object[] {
                                                               "1bpp (Black and White)",
                                                               "8bpp (256 colors optimized palette)",
                                                               "24bpp (true color)"});
            this.cbPixelFormat.Location = new System.Drawing.Point(16, 40);
            this.cbPixelFormat.Name = "cbPixelFormat";
            this.cbPixelFormat.Size = new System.Drawing.Size(208, 21);
            this.cbPixelFormat.TabIndex = 3;
            // 
            // btnOk
            // 
            this.btnOk.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.btnOk.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.btnOk.Location = new System.Drawing.Point(148, 168);
            this.btnOk.Name = "btnOk";
            this.btnOk.TabIndex = 4;
            this.btnOk.Text = "Ok";
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.btnCancel.Location = new System.Drawing.Point(244, 168);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.TabIndex = 5;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.cbTransparent);
            this.panel1.Controls.Add(this.edPercent);
            this.panel1.Controls.Add(this.cbPixelFormat);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Location = new System.Drawing.Point(16, 16);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(240, 136);
            this.panel1.TabIndex = 6;
            // 
            // panel2
            // 
            this.panel2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                | System.Windows.Forms.AnchorStyles.Left)
                | System.Windows.Forms.AnchorStyles.Right)));
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel2.Controls.Add(this.pictureBox1);
            this.panel2.Location = new System.Drawing.Point(280, 16);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(168, 136);
            this.panel2.TabIndex = 7;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pictureBox1.Location = new System.Drawing.Point(0, 0);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(164, 132);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // cbTransparent
            // 
            this.cbTransparent.Location = new System.Drawing.Point(16, 88);
            this.cbTransparent.Name = "cbTransparent";
            this.cbTransparent.TabIndex = 4;
            this.cbTransparent.Text = "Transparent";
            // 
            // TCompressForm
            // 
            this.AcceptButton = this.btnOk;
            this.CancelButton = this.btnCancel;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(472, 214);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOk);
            this.Name = "TCompressForm";
            this.Text = "Compression Options...";
            this.Load += new System.EventHandler(this.TCompressForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.edPercent)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);

        }
        #endregion
    }
}

