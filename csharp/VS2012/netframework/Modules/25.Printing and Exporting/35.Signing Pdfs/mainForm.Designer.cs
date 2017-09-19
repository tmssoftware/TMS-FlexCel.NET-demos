namespace SigningPdfs
{
    partial class mainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
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
            this.btnCreateAndSign = new System.Windows.Forms.Button();
            this.cbVisibleSignature = new System.Windows.Forms.CheckBox();
            this.OpenExcelDialog = new System.Windows.Forms.OpenFileDialog();
            this.savePdfDialog = new System.Windows.Forms.SaveFileDialog();
            this.SignaturePicture = new System.Windows.Forms.PictureBox();
            this.OpenImageDialog = new System.Windows.Forms.OpenFileDialog();
            ((System.ComponentModel.ISupportInitialize)(this.SignaturePicture)).BeginInit();
            this.SuspendLayout();
            // 
            // btnCreateAndSign
            // 
            this.btnCreateAndSign.Image = global::SigningPdfs.Properties.Resources.acroread;
            this.btnCreateAndSign.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnCreateAndSign.Location = new System.Drawing.Point(24, 25);
            this.btnCreateAndSign.Name = "btnCreateAndSign";
            this.btnCreateAndSign.Size = new System.Drawing.Size(155, 30);
            this.btnCreateAndSign.TabIndex = 0;
            this.btnCreateAndSign.Text = "Create and Sign Pdf";
            this.btnCreateAndSign.UseVisualStyleBackColor = true;
            this.btnCreateAndSign.Click += new System.EventHandler(this.btnCreateAndSign_Click);
            // 
            // cbVisibleSignature
            // 
            this.cbVisibleSignature.AutoSize = true;
            this.cbVisibleSignature.Location = new System.Drawing.Point(24, 78);
            this.cbVisibleSignature.Name = "cbVisibleSignature";
            this.cbVisibleSignature.Size = new System.Drawing.Size(167, 17);
            this.cbVisibleSignature.TabIndex = 1;
            this.cbVisibleSignature.Text = "Visible Signature (in last page)";
            this.cbVisibleSignature.UseVisualStyleBackColor = true;
            this.cbVisibleSignature.CheckedChanged += new System.EventHandler(this.cbVisibleSignature_CheckedChanged);
            // 
            // OpenExcelDialog
            // 
            this.OpenExcelDialog.DefaultExt = "xls";
            this.OpenExcelDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|Excel 97/2003|*.xls|Excel 2007|*.xlsx;*.xlsm|All files|*.*";
            this.OpenExcelDialog.Title = "Select Excel file to convert...";
            // 
            // savePdfDialog
            // 
            this.savePdfDialog.DefaultExt = "pdf";
            this.savePdfDialog.Filter = "Pdf Files|*.pdf";
            this.savePdfDialog.Title = "Select where to save the file...";
            // 
            // SignaturePicture
            // 
            this.SignaturePicture.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.SignaturePicture.Image = global::SigningPdfs.Properties.Resources.sign;
            this.SignaturePicture.Location = new System.Drawing.Point(24, 110);
            this.SignaturePicture.Name = "SignaturePicture";
            this.SignaturePicture.Size = new System.Drawing.Size(155, 100);
            this.SignaturePicture.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.SignaturePicture.TabIndex = 2;
            this.SignaturePicture.TabStop = false;
            this.SignaturePicture.Click += new System.EventHandler(this.SignaturePicture_Click);
            // 
            // OpenImageDialog
            // 
            this.OpenImageDialog.Filter = "Supported Images|*.png;*.bmp*.jpg|All files|*.*";
            // 
            // mainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(205, 100);
            this.Controls.Add(this.SignaturePicture);
            this.Controls.Add(this.cbVisibleSignature);
            this.Controls.Add(this.btnCreateAndSign);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Name = "mainForm";
            this.Text = "Signing PDFs";
            ((System.ComponentModel.ISupportInitialize)(this.SignaturePicture)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCreateAndSign;
        private System.Windows.Forms.CheckBox cbVisibleSignature;
        private System.Windows.Forms.OpenFileDialog OpenExcelDialog;
        private System.Windows.Forms.SaveFileDialog savePdfDialog;
        private System.Windows.Forms.PictureBox SignaturePicture;
        private System.Windows.Forms.OpenFileDialog OpenImageDialog;
    }
}

