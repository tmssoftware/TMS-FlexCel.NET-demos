using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Text;
using FlexCel.Core;
using FlexCel.XlsAdapter;
namespace CopyAndPaste
{
    public partial class mainForm: System.Windows.Forms.Form
    {
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
            this.btnPaste = new System.Windows.Forms.Button();
            this.btnNewFile = new System.Windows.Forms.Button();
            this.btnCopy = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.btnDragMe = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.DropHere = new System.Windows.Forms.Label();
            this.btnOpenFile = new System.Windows.Forms.Button();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.SuspendLayout();
            // 
            // btnPaste
            // 
            this.btnPaste.Location = new System.Drawing.Point(44, 251);
            this.btnPaste.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.btnPaste.Name = "btnPaste";
            this.btnPaste.Size = new System.Drawing.Size(138, 42);
            this.btnPaste.TabIndex = 0;
            this.btnPaste.Text = "Paste";
            this.btnPaste.Click += new System.EventHandler(this.btnPaste_Click);
            // 
            // btnNewFile
            // 
            this.btnNewFile.Location = new System.Drawing.Point(44, 74);
            this.btnNewFile.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.btnNewFile.Name = "btnNewFile";
            this.btnNewFile.Size = new System.Drawing.Size(138, 42);
            this.btnNewFile.TabIndex = 1;
            this.btnNewFile.Text = "New File";
            this.btnNewFile.Click += new System.EventHandler(this.btnNewFile_Click);
            // 
            // btnCopy
            // 
            this.btnCopy.Location = new System.Drawing.Point(44, 428);
            this.btnCopy.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.btnCopy.Name = "btnCopy";
            this.btnCopy.Size = new System.Drawing.Size(138, 42);
            this.btnCopy.TabIndex = 2;
            this.btnCopy.Text = "Copy";
            this.btnCopy.Click += new System.EventHandler(this.btnCopy_Click);
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.LightSkyBlue;
            this.label1.Location = new System.Drawing.Point(44, 15);
            this.label1.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(719, 44);
            this.label1.TabIndex = 4;
            this.label1.Text = "1) Begin by creating a new file or opening an existing file...";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.Color.LightSkyBlue;
            this.label2.Location = new System.Drawing.Point(44, 148);
            this.label2.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(719, 44);
            this.label2.TabIndex = 5;
            this.label2.Text = "2) Now go to Excel, copy some cells and paste them here...";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.Color.LightSkyBlue;
            this.label3.Location = new System.Drawing.Point(44, 325);
            this.label3.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(719, 44);
            this.label3.TabIndex = 6;
            this.label3.Text = "3) After pasting, you can copy back the results to the clipboard";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label4
            // 
            this.label4.BackColor = System.Drawing.Color.SteelBlue;
            this.label4.ForeColor = System.Drawing.Color.White;
            this.label4.Location = new System.Drawing.Point(44, 369);
            this.label4.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(719, 44);
            this.label4.TabIndex = 7;
            this.label4.Text = "Press the \"Copy\" button or drag the \"Drag Me!\" into Excel.";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnDragMe
            // 
            this.btnDragMe.AllowDrop = true;
            this.btnDragMe.Location = new System.Drawing.Point(205, 428);
            this.btnDragMe.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.btnDragMe.Name = "btnDragMe";
            this.btnDragMe.Size = new System.Drawing.Size(138, 42);
            this.btnDragMe.TabIndex = 8;
            this.btnDragMe.Text = "Drag Me!";
            this.btnDragMe.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btnDragMe_MouseDown);
            // 
            // label5
            // 
            this.label5.BackColor = System.Drawing.Color.SteelBlue;
            this.label5.ForeColor = System.Drawing.Color.White;
            this.label5.Location = new System.Drawing.Point(44, 192);
            this.label5.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(719, 44);
            this.label5.TabIndex = 10;
            this.label5.Text = "Press the \"Paste\" button or drag some cells from Excel into \"Drop Here!\".";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // DropHere
            // 
            this.DropHere.AllowDrop = true;
            this.DropHere.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.DropHere.Location = new System.Drawing.Point(200, 251);
            this.DropHere.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.DropHere.Name = "DropHere";
            this.DropHere.Size = new System.Drawing.Size(183, 42);
            this.DropHere.TabIndex = 11;
            this.DropHere.Text = "Drop Here!";
            this.DropHere.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.DropHere.DragDrop += new System.Windows.Forms.DragEventHandler(this.DropHere_DragDrop);
            this.DropHere.DragOver += new System.Windows.Forms.DragEventHandler(this.DropHere_DragOver);
            // 
            // btnOpenFile
            // 
            this.btnOpenFile.Location = new System.Drawing.Point(205, 74);
            this.btnOpenFile.Margin = new System.Windows.Forms.Padding(6);
            this.btnOpenFile.Name = "btnOpenFile";
            this.btnOpenFile.Size = new System.Drawing.Size(138, 42);
            this.btnOpenFile.TabIndex = 12;
            this.btnOpenFile.Text = "Open File";
            this.btnOpenFile.Click += new System.EventHandler(this.btnOpenFile_Click);
            // 
            // openFileDialog
            // 
            this.openFileDialog.DefaultExt = "xls";
            this.openFileDialog.Filter = "Excel Files|*.xls|All files|*.*";
            this.openFileDialog.Title = "Select a file to preview";
            // 
            // mainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(836, 615);
            this.Controls.Add(this.btnOpenFile);
            this.Controls.Add(this.DropHere);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.btnDragMe);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnCopy);
            this.Controls.Add(this.btnNewFile);
            this.Controls.Add(this.btnPaste);
            this.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.Name = "mainForm";
            this.Text = "Copy and Paste Demo";
            this.ResumeLayout(false);

        }
        #endregion

        private Button btnPaste;
        private Button btnNewFile;
        private Button btnCopy;
        private Label label1;
        private Label label2;
        private Label label3;
        private Label label4;
        private Label label5;
        private Button btnDragMe;
        private Label DropHere;
        private Button btnOpenFile;
        private OpenFileDialog openFileDialog;

    }
}

