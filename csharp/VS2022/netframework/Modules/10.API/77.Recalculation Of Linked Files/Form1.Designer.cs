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
using System.Globalization;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
namespace RecalculationOfLinkedFiles
{
    public partial class mainForm: System.Windows.Forms.Form
    {
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TextBox CellA1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox Cell2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox Cell3;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox Cell4;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.TextBox ChartA1;
        private System.Windows.Forms.TextBox ChartB1;
        private System.Windows.Forms.TextBox ChartB2;
        private System.Windows.Forms.TextBox ChartA2;
        private System.Windows.Forms.TextBox ChartB3;
        private System.Windows.Forms.TextBox ChartA3;
        private System.Windows.Forms.PictureBox chartBox;
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
            this.panel2 = new System.Windows.Forms.Panel();
            this.button2 = new System.Windows.Forms.Button();
            this.CellA1 = new System.Windows.Forms.TextBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.Cell4 = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.Cell3 = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.Cell2 = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label18 = new System.Windows.Forms.Label();
            this.ChartA1 = new System.Windows.Forms.TextBox();
            this.panel3 = new System.Windows.Forms.Panel();
            this.chartBox = new System.Windows.Forms.PictureBox();
            this.ChartB3 = new System.Windows.Forms.TextBox();
            this.ChartA3 = new System.Windows.Forms.TextBox();
            this.ChartB2 = new System.Windows.Forms.TextBox();
            this.ChartA2 = new System.Windows.Forms.TextBox();
            this.ChartB1 = new System.Windows.Forms.TextBox();
            this.panel2.SuspendLayout();
            this.panel1.SuspendLayout();
            this.panel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.chartBox)).BeginInit();
            this.SuspendLayout();
            // 
            // panel2
            // 
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel2.Controls.Add(this.button2);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel2.Location = new System.Drawing.Point(0, 422);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(760, 32);
            this.panel2.TabIndex = 2;
            // 
            // button2
            // 
            this.button2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button2.BackColor = System.Drawing.SystemColors.Control;
            this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
            this.button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button2.Location = new System.Drawing.Point(697, 2);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(56, 26);
            this.button2.TabIndex = 2;
            this.button2.Text = "Exit";
            this.button2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // CellA1
            // 
            this.CellA1.Location = new System.Drawing.Point(24, 88);
            this.CellA1.Name = "CellA1";
            this.CellA1.Size = new System.Drawing.Size(100, 20);
            this.CellA1.TabIndex = 3;
            this.CellA1.TextChanged += new System.EventHandler(this.CellA1_TextChanged);
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.label8);
            this.panel1.Controls.Add(this.label9);
            this.panel1.Controls.Add(this.Cell4);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.label7);
            this.panel1.Controls.Add(this.Cell3);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.Cell2);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.CellA1);
            this.panel1.Location = new System.Drawing.Point(8, 8);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(736, 136);
            this.panel1.TabIndex = 4;
            // 
            // label8
            // 
            this.label8.Location = new System.Drawing.Point(520, 72);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(160, 16);
            this.label8.TabIndex = 15;
            this.label8.Text = "=[Third File.xls]Sheet1!A1 + 7";
            // 
            // label9
            // 
            this.label9.Location = new System.Drawing.Point(520, 56);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(100, 16);
            this.label9.TabIndex = 14;
            this.label9.Text = "First File: A2";
            // 
            // Cell4
            // 
            this.Cell4.Enabled = false;
            this.Cell4.Location = new System.Drawing.Point(520, 88);
            this.Cell4.Name = "Cell4";
            this.Cell4.Size = new System.Drawing.Size(152, 20);
            this.Cell4.TabIndex = 13;
            // 
            // label6
            // 
            this.label6.Location = new System.Drawing.Point(328, 72);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(184, 16);
            this.label6.TabIndex = 12;
            this.label6.Text = "=[Second File.xls]Sheet1!A1 * 5";
            // 
            // label7
            // 
            this.label7.Location = new System.Drawing.Point(328, 56);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(100, 16);
            this.label7.TabIndex = 11;
            this.label7.Text = "Third File: A1";
            // 
            // Cell3
            // 
            this.Cell3.Enabled = false;
            this.Cell3.Location = new System.Drawing.Point(328, 88);
            this.Cell3.Name = "Cell3";
            this.Cell3.Size = new System.Drawing.Size(152, 20);
            this.Cell3.TabIndex = 10;
            // 
            // label5
            // 
            this.label5.Location = new System.Drawing.Point(152, 72);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(152, 16);
            this.label5.TabIndex = 9;
            this.label5.Text = "=[First File.xls]Sheet1!A1 * 2";
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(24, 72);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(100, 16);
            this.label4.TabIndex = 8;
            this.label4.Text = "Constant";
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(152, 56);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(100, 16);
            this.label3.TabIndex = 7;
            this.label3.Text = "Second File: A1";
            // 
            // Cell2
            // 
            this.Cell2.Enabled = false;
            this.Cell2.Location = new System.Drawing.Point(152, 88);
            this.Cell2.Name = "Cell2";
            this.Cell2.Size = new System.Drawing.Size(144, 20);
            this.Cell2.TabIndex = 6;
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(24, 56);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(100, 16);
            this.label2.TabIndex = 5;
            this.label2.Text = "First File: A1";
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.Location = new System.Drawing.Point(16, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(704, 32);
            this.label1.TabIndex = 4;
            this.label1.Text = "In this first example we will dynamically create 3 linked files. We will create a" +
    " workspace to link the files, and see how recalculation works.";
            // 
            // label18
            // 
            this.label18.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label18.Location = new System.Drawing.Point(16, 16);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(704, 32);
            this.label18.TabIndex = 4;
            this.label18.Text = "This second example shows how to load files when we don\'t know a priori which fil" +
    "es we need to recalculate a file. To make it more interesting, we will use a cha" +
    "rt linked to other file.";
            // 
            // ChartA1
            // 
            this.ChartA1.Location = new System.Drawing.Point(16, 56);
            this.ChartA1.Name = "ChartA1";
            this.ChartA1.Size = new System.Drawing.Size(64, 20);
            this.ChartA1.TabIndex = 3;
            this.ChartA1.Text = "1";
            this.ChartA1.TextChanged += new System.EventHandler(this.Chart_TextChanged);
            // 
            // panel3
            // 
            this.panel3.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel3.Controls.Add(this.chartBox);
            this.panel3.Controls.Add(this.ChartB3);
            this.panel3.Controls.Add(this.ChartA3);
            this.panel3.Controls.Add(this.ChartB2);
            this.panel3.Controls.Add(this.ChartA2);
            this.panel3.Controls.Add(this.ChartB1);
            this.panel3.Controls.Add(this.label18);
            this.panel3.Controls.Add(this.ChartA1);
            this.panel3.Location = new System.Drawing.Point(8, 176);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(736, 225);
            this.panel3.TabIndex = 5;
            // 
            // chartBox
            // 
            this.chartBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.chartBox.Location = new System.Drawing.Point(176, 56);
            this.chartBox.Name = "chartBox";
            this.chartBox.Size = new System.Drawing.Size(544, 152);
            this.chartBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.chartBox.TabIndex = 52;
            this.chartBox.TabStop = false;
            // 
            // ChartB3
            // 
            this.ChartB3.Location = new System.Drawing.Point(88, 104);
            this.ChartB3.Name = "ChartB3";
            this.ChartB3.Size = new System.Drawing.Size(64, 20);
            this.ChartB3.TabIndex = 9;
            this.ChartB3.Text = "5";
            this.ChartB3.TextChanged += new System.EventHandler(this.Chart_TextChanged);
            // 
            // ChartA3
            // 
            this.ChartA3.Location = new System.Drawing.Point(16, 104);
            this.ChartA3.Name = "ChartA3";
            this.ChartA3.Size = new System.Drawing.Size(64, 20);
            this.ChartA3.TabIndex = 8;
            this.ChartA3.Text = "3";
            this.ChartA3.TextChanged += new System.EventHandler(this.Chart_TextChanged);
            // 
            // ChartB2
            // 
            this.ChartB2.Location = new System.Drawing.Point(88, 80);
            this.ChartB2.Name = "ChartB2";
            this.ChartB2.Size = new System.Drawing.Size(64, 20);
            this.ChartB2.TabIndex = 7;
            this.ChartB2.Text = "4";
            this.ChartB2.TextChanged += new System.EventHandler(this.Chart_TextChanged);
            // 
            // ChartA2
            // 
            this.ChartA2.Location = new System.Drawing.Point(16, 80);
            this.ChartA2.Name = "ChartA2";
            this.ChartA2.Size = new System.Drawing.Size(64, 20);
            this.ChartA2.TabIndex = 6;
            this.ChartA2.Text = "2";
            this.ChartA2.TextChanged += new System.EventHandler(this.Chart_TextChanged);
            // 
            // ChartB1
            // 
            this.ChartB1.Location = new System.Drawing.Point(88, 56);
            this.ChartB1.Name = "ChartB1";
            this.ChartB1.Size = new System.Drawing.Size(64, 20);
            this.ChartB1.TabIndex = 5;
            this.ChartB1.Text = "3";
            this.ChartB1.TextChanged += new System.EventHandler(this.Chart_TextChanged);
            // 
            // mainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(760, 454);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.panel2);
            this.Name = "mainForm";
            this.Text = "Calculation of linked files";
            this.panel2.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.chartBox)).EndInit();
            this.ResumeLayout(false);

        }
        #endregion
    }
}

