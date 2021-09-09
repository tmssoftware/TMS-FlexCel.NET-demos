using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using FlexCel.Core;
using System.Text;
namespace ExportHTML
{
    public partial class Mailform: System.Windows.Forms.Form
    {


        private System.Windows.Forms.Button btnEmail;
        private System.Windows.Forms.TextBox edOutServer;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox edTo;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox edFrom;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox edSubject;
        private System.Windows.Forms.Label label5;
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Mailform));
            this.btnEmail = new System.Windows.Forms.Button();
            this.edOutServer = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.edTo = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.edFrom = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.edSubject = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnEmail
            // 
            this.btnEmail.BackColor = System.Drawing.SystemColors.Control;
            this.btnEmail.Image = ((System.Drawing.Image)(resources.GetObject("btnEmail.Image")));
            this.btnEmail.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnEmail.Location = new System.Drawing.Point(200, 200);
            this.btnEmail.Name = "btnEmail";
            this.btnEmail.Size = new System.Drawing.Size(70, 30);
            this.btnEmail.TabIndex = 4;
            this.btnEmail.Text = "e-mail!";
            this.btnEmail.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnEmail.UseVisualStyleBackColor = false;
            this.btnEmail.Click += new System.EventHandler(this.btnEmail_Click);
            // 
            // edOutServer
            // 
            this.edOutServer.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.edOutServer.Location = new System.Drawing.Point(136, 144);
            this.edOutServer.Name = "edOutServer";
            this.edOutServer.Size = new System.Drawing.Size(304, 20);
            this.edOutServer.TabIndex = 3;
            this.edOutServer.Text = "pop.mycompany.com";
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(8, 152);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(128, 16);
            this.label2.TabIndex = 10;
            this.label2.Text = "Outgoing Mail Server:";
            // 
            // edTo
            // 
            this.edTo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.edTo.Location = new System.Drawing.Point(136, 40);
            this.edTo.Name = "edTo";
            this.edTo.Size = new System.Drawing.Size(304, 20);
            this.edTo.TabIndex = 1;
            this.edTo.Text = "user@hiscompany.com";
            this.edTo.Leave += new System.EventHandler(this.edTo_Leave);
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(16, 40);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(128, 16);
            this.label3.TabIndex = 14;
            this.label3.Text = "To:";
            // 
            // edFrom
            // 
            this.edFrom.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.edFrom.Location = new System.Drawing.Point(136, 8);
            this.edFrom.Name = "edFrom";
            this.edFrom.Size = new System.Drawing.Size(304, 20);
            this.edFrom.TabIndex = 0;
            this.edFrom.Text = "myname@mycompany.com";
            this.edFrom.Leave += new System.EventHandler(this.edFrom_Leave);
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(16, 16);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(120, 16);
            this.label4.TabIndex = 12;
            this.label4.Text = "From:";
            // 
            // edSubject
            // 
            this.edSubject.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.edSubject.Location = new System.Drawing.Point(136, 80);
            this.edSubject.Name = "edSubject";
            this.edSubject.Size = new System.Drawing.Size(304, 20);
            this.edSubject.TabIndex = 2;
            this.edSubject.Text = "A test from FlexCel";
            // 
            // label5
            // 
            this.label5.Location = new System.Drawing.Point(16, 80);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(128, 16);
            this.label5.TabIndex = 16;
            this.label5.Text = "Subject:";
            // 
            // Mailform
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(472, 246);
            this.Controls.Add(this.edSubject);
            this.Controls.Add(this.edTo);
            this.Controls.Add(this.edFrom);
            this.Controls.Add(this.edOutServer);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnEmail);
            this.Name = "Mailform";
            this.ShowInTaskbar = false;
            this.Text = "Send email...";
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion
    }
}

