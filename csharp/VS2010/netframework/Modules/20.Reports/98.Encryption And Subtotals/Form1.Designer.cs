using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using FlexCel.Report;
using FlexCel.Demo.SharedData;
namespace EncryptionAndSubtotals
{
    public partial class mainForm: System.Windows.Forms.Form
    {
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox OpenPassTemplate;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox OpenPassGenerated;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox ModifyPassGenerated;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox ProtectWorkbookPass;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox ProtectSheetPass;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ComboBox encryptionType;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox ReservingUser;
        private System.Windows.Forms.CheckBox RecommendReadOnly;
        private System.Windows.Forms.CheckBox ProtectWorkbook;
        private System.Windows.Forms.CheckBox ProtectSheet;
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
            this.button1 = new System.Windows.Forms.Button();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.btnCancel = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.OpenPassTemplate = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.ProtectSheet = new System.Windows.Forms.CheckBox();
            this.ProtectWorkbook = new System.Windows.Forms.CheckBox();
            this.RecommendReadOnly = new System.Windows.Forms.CheckBox();
            this.label9 = new System.Windows.Forms.Label();
            this.ReservingUser = new System.Windows.Forms.TextBox();
            this.encryptionType = new System.Windows.Forms.ComboBox();
            this.label8 = new System.Windows.Forms.Label();
            this.ProtectSheetPass = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.ProtectWorkbookPass = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.ModifyPassGenerated = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.OpenPassGenerated = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.BackColor = System.Drawing.Color.Green;
            this.button1.ForeColor = System.Drawing.Color.White;
            this.button1.Location = new System.Drawing.Point(312, 414);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(112, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "GO!";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|Excel 97/2003|*.xls|Excel 2007|*.xlsx;*.xlsm|All " +
    "files|*.*";
            this.saveFileDialog1.RestoreDirectory = true;
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.ForeColor = System.Drawing.Color.White;
            this.btnCancel.Location = new System.Drawing.Point(432, 414);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(112, 23);
            this.btnCancel.TabIndex = 3;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.OpenPassTemplate);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Location = new System.Drawing.Point(24, 16);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(496, 80);
            this.panel1.TabIndex = 6;
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(8, 32);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(200, 16);
            this.label2.TabIndex = 8;
            this.label2.Text = "Password to open the template:";
            // 
            // OpenPassTemplate
            // 
            this.OpenPassTemplate.Location = new System.Drawing.Point(8, 48);
            this.OpenPassTemplate.Name = "OpenPassTemplate";
            this.OpenPassTemplate.Size = new System.Drawing.Size(200, 20);
            this.OpenPassTemplate.TabIndex = 7;
            this.OpenPassTemplate.Text = "flexcel";
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(8, 8);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(416, 24);
            this.label1.TabIndex = 6;
            this.label1.Text = "The template is protected with a password to open. On this demo, it is \"flexcel\"";
            // 
            // panel2
            // 
            this.panel2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel2.Controls.Add(this.ProtectSheet);
            this.panel2.Controls.Add(this.ProtectWorkbook);
            this.panel2.Controls.Add(this.RecommendReadOnly);
            this.panel2.Controls.Add(this.label9);
            this.panel2.Controls.Add(this.ReservingUser);
            this.panel2.Controls.Add(this.encryptionType);
            this.panel2.Controls.Add(this.label8);
            this.panel2.Controls.Add(this.ProtectSheetPass);
            this.panel2.Controls.Add(this.label7);
            this.panel2.Controls.Add(this.label6);
            this.panel2.Controls.Add(this.ProtectWorkbookPass);
            this.panel2.Controls.Add(this.label5);
            this.panel2.Controls.Add(this.ModifyPassGenerated);
            this.panel2.Controls.Add(this.label3);
            this.panel2.Controls.Add(this.OpenPassGenerated);
            this.panel2.Controls.Add(this.label4);
            this.panel2.Location = new System.Drawing.Point(24, 112);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(496, 296);
            this.panel2.TabIndex = 7;
            // 
            // ProtectSheet
            // 
            this.ProtectSheet.Checked = true;
            this.ProtectSheet.CheckState = System.Windows.Forms.CheckState.Checked;
            this.ProtectSheet.Location = new System.Drawing.Point(368, 203);
            this.ProtectSheet.Name = "ProtectSheet";
            this.ProtectSheet.Size = new System.Drawing.Size(64, 16);
            this.ProtectSheet.TabIndex = 22;
            this.ProtectSheet.Text = "Protect";
            // 
            // ProtectWorkbook
            // 
            this.ProtectWorkbook.Checked = true;
            this.ProtectWorkbook.CheckState = System.Windows.Forms.CheckState.Checked;
            this.ProtectWorkbook.Location = new System.Drawing.Point(368, 147);
            this.ProtectWorkbook.Name = "ProtectWorkbook";
            this.ProtectWorkbook.Size = new System.Drawing.Size(64, 16);
            this.ProtectWorkbook.TabIndex = 21;
            this.ProtectWorkbook.Text = "Protect";
            // 
            // RecommendReadOnly
            // 
            this.RecommendReadOnly.Location = new System.Drawing.Point(240, 243);
            this.RecommendReadOnly.Name = "RecommendReadOnly";
            this.RecommendReadOnly.Size = new System.Drawing.Size(168, 24);
            this.RecommendReadOnly.TabIndex = 20;
            this.RecommendReadOnly.Text = "Recommend read only";
            // 
            // label9
            // 
            this.label9.Location = new System.Drawing.Point(8, 227);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(200, 16);
            this.label9.TabIndex = 19;
            this.label9.Text = "Reserving user: (for modify password)";
            // 
            // ReservingUser
            // 
            this.ReservingUser.Location = new System.Drawing.Point(8, 243);
            this.ReservingUser.Name = "ReservingUser";
            this.ReservingUser.Size = new System.Drawing.Size(200, 20);
            this.ReservingUser.TabIndex = 18;
            this.ReservingUser.Text = "Flexcel User";
            // 
            // encryptionType
            // 
            this.encryptionType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.encryptionType.Items.AddRange(new object[] {
            "Default Excel 97/2000 Encryption",
            "Excel 95 XOR Encryption"});
            this.encryptionType.Location = new System.Drawing.Point(8, 84);
            this.encryptionType.Name = "encryptionType";
            this.encryptionType.Size = new System.Drawing.Size(424, 21);
            this.encryptionType.TabIndex = 17;
            // 
            // label8
            // 
            this.label8.Location = new System.Drawing.Point(8, 48);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(488, 33);
            this.label8.TabIndex = 16;
            this.label8.Text = "Encryption type for xls files (xlsx uses Agile encryption). Note that this is onl" +
    "y needed when saving, as the encryption type is autodetected when opening:";
            // 
            // ProtectSheetPass
            // 
            this.ProtectSheetPass.Location = new System.Drawing.Point(232, 176);
            this.ProtectSheetPass.Name = "ProtectSheetPass";
            this.ProtectSheetPass.Size = new System.Drawing.Size(120, 20);
            this.ProtectSheetPass.TabIndex = 15;
            this.ProtectSheetPass.Text = "sheet";
            // 
            // label7
            // 
            this.label7.Location = new System.Drawing.Point(232, 179);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(248, 16);
            this.label7.TabIndex = 14;
            this.label7.Text = "Password to protect the generated sheets:";
            // 
            // label6
            // 
            this.label6.Location = new System.Drawing.Point(232, 131);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(248, 16);
            this.label6.TabIndex = 12;
            this.label6.Text = "Password to protect the generated workbook:";
            // 
            // ProtectWorkbookPass
            // 
            this.ProtectWorkbookPass.Location = new System.Drawing.Point(232, 147);
            this.ProtectWorkbookPass.Name = "ProtectWorkbookPass";
            this.ProtectWorkbookPass.Size = new System.Drawing.Size(120, 20);
            this.ProtectWorkbookPass.TabIndex = 11;
            this.ProtectWorkbookPass.Text = "workbook";
            // 
            // label5
            // 
            this.label5.Location = new System.Drawing.Point(8, 179);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(200, 16);
            this.label5.TabIndex = 10;
            this.label5.Text = "Password to modify the generated file:";
            // 
            // ModifyPassGenerated
            // 
            this.ModifyPassGenerated.Location = new System.Drawing.Point(8, 195);
            this.ModifyPassGenerated.Name = "ModifyPassGenerated";
            this.ModifyPassGenerated.Size = new System.Drawing.Size(200, 20);
            this.ModifyPassGenerated.TabIndex = 9;
            this.ModifyPassGenerated.Text = "modify";
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(8, 131);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(200, 16);
            this.label3.TabIndex = 8;
            this.label3.Text = "Password to open the generated file:";
            // 
            // OpenPassGenerated
            // 
            this.OpenPassGenerated.Location = new System.Drawing.Point(8, 147);
            this.OpenPassGenerated.Name = "OpenPassGenerated";
            this.OpenPassGenerated.Size = new System.Drawing.Size(200, 20);
            this.OpenPassGenerated.TabIndex = 7;
            this.OpenPassGenerated.Text = "open";
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(8, 8);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(416, 40);
            this.label4.TabIndex = 6;
            this.label4.Text = "Here we enter the passwords we want to protect the generated sheets and workbook." +
    " Leave them blank to have no password.";
            // 
            // mainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(552, 443);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.button1);
            this.Name = "mainForm";
            this.Text = "Encryption And Subtotals";
            this.Load += new System.EventHandler(this.mainForm_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.ResumeLayout(false);

        }
        #endregion
    }
}

