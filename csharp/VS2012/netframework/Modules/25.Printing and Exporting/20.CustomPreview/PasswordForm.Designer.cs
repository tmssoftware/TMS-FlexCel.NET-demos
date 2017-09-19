using System;
using System.Drawing;
using System.ComponentModel;
using System.Windows.Forms;
namespace CustomPreview
{
    public partial class PasswordForm: System.Windows.Forms.Form
    {
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox PasswordEdit;
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
            this.btnOk = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.PasswordEdit = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(24, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(240, 23);
            this.label1.TabIndex = 0;
            this.label1.Text = "Please enter the password to open this file:";
            // 
            // btnOk
            // 
            this.btnOk.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnOk.Location = new System.Drawing.Point(152, 112);
            this.btnOk.Name = "btnOk";
            this.btnOk.TabIndex = 1;
            this.btnOk.Text = "Ok";
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(40, 64);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(64, 23);
            this.label3.TabIndex = 5;
            this.label3.Text = "Password:";
            // 
            // PasswordEdit
            // 
            this.PasswordEdit.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                | System.Windows.Forms.AnchorStyles.Right)));
            this.PasswordEdit.Location = new System.Drawing.Point(112, 64);
            this.PasswordEdit.Name = "PasswordEdit";
            this.PasswordEdit.PasswordChar = '*';
            this.PasswordEdit.Size = new System.Drawing.Size(200, 20);
            this.PasswordEdit.TabIndex = 0;
            this.PasswordEdit.Text = "";
            // 
            // PasswordForm
            // 
            this.AcceptButton = this.btnOk;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(408, 154);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.PasswordEdit);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.label1);
            this.Name = "PasswordForm";
            this.ShowInTaskbar = false;
            this.Text = "File is password protected.";
            this.ResumeLayout(false);

        }
        #endregion
    }
}

