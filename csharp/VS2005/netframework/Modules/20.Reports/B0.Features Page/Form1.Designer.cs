using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Data.OleDb;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using FlexCel.Report;
using FlexCel.Render;
using FlexCel.Pdf;
using System.Globalization;
using System.Xml;
namespace FeaturesPage
{
    public partial class mainForm: System.Windows.Forms.Form
    {
        private System.Windows.Forms.Label label1;
        private System.Data.OleDb.OleDbCommand oleDbSelectCommand1;
        private System.Data.OleDb.OleDbCommand oleDbInsertCommand1;
        private System.Data.OleDb.OleDbCommand oleDbUpdateCommand1;
        private System.Data.OleDb.OleDbCommand oleDbDeleteCommand1;
        private System.Data.OleDb.OleDbConnection dbconnection;
        private System.Data.OleDb.OleDbDataAdapter categoriesAdapter;
        private System.Data.OleDb.OleDbDataAdapter featuresAdapter;
        private System.Data.OleDb.OleDbCommand oleDbSelectCommand2;
        private System.Data.OleDb.OleDbCommand oleDbInsertCommand2;
        private System.Data.OleDb.OleDbCommand oleDbUpdateCommand2;
        private System.Data.OleDb.OleDbCommand oleDbDeleteCommand2;
        private System.Data.OleDb.OleDbCommand oleDbSelectCommand3;
        private System.Data.OleDb.OleDbDataAdapter hyperlinksAdapter;
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
            this.label1 = new System.Windows.Forms.Label();
            this.oleDbSelectCommand1 = new System.Data.OleDb.OleDbCommand();
            this.dbconnection = new System.Data.OleDb.OleDbConnection();
            this.oleDbInsertCommand1 = new System.Data.OleDb.OleDbCommand();
            this.oleDbUpdateCommand1 = new System.Data.OleDb.OleDbCommand();
            this.oleDbDeleteCommand1 = new System.Data.OleDb.OleDbCommand();
            this.categoriesAdapter = new System.Data.OleDb.OleDbDataAdapter();
            this.featuresAdapter = new System.Data.OleDb.OleDbDataAdapter();
            this.oleDbDeleteCommand2 = new System.Data.OleDb.OleDbCommand();
            this.oleDbInsertCommand2 = new System.Data.OleDb.OleDbCommand();
            this.oleDbSelectCommand2 = new System.Data.OleDb.OleDbCommand();
            this.oleDbUpdateCommand2 = new System.Data.OleDb.OleDbCommand();
            this.oleDbSelectCommand3 = new System.Data.OleDb.OleDbCommand();
            this.hyperlinksAdapter = new System.Data.OleDb.OleDbDataAdapter();
            this.mainToolbar = new System.Windows.Forms.ToolStrip();
            this.toolStripButton4 = new System.Windows.Forms.ToolStripButton();
            this.toolStripButton3 = new System.Windows.Forms.ToolStripButton();
            this.toolStripButton2 = new System.Windows.Forms.ToolStripButton();
            this.toolStripButton1 = new System.Windows.Forms.ToolStripButton();
            this.mainToolbar.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(179, 69);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(200, 48);
            this.label1.TabIndex = 4;
            this.label1.Text = "Files will be saved under a \"Features\" Folder in under where the application is r" +
    "unning.";
            // 
            // oleDbSelectCommand1
            // 
            this.oleDbSelectCommand1.CommandText = "SELECT Caption, CategoryId, CategoryName, Description FROM Categories";
            this.oleDbSelectCommand1.Connection = this.dbconnection;
            // 
            // dbconnection
            // 
            this.dbconnection.ConnectionString = resources.GetString("dbconnection.ConnectionString");
            // 
            // oleDbInsertCommand1
            // 
            this.oleDbInsertCommand1.CommandText = "INSERT INTO Categories(Caption, CategoryName, Description) VALUES (?, ?, ?)";
            this.oleDbInsertCommand1.Connection = this.dbconnection;
            this.oleDbInsertCommand1.Parameters.AddRange(new System.Data.OleDb.OleDbParameter[] {
            new System.Data.OleDb.OleDbParameter("Caption", System.Data.OleDb.OleDbType.VarWChar, 255, "Caption"),
            new System.Data.OleDb.OleDbParameter("CategoryName", System.Data.OleDb.OleDbType.VarWChar, 255, "CategoryName"),
            new System.Data.OleDb.OleDbParameter("Description", System.Data.OleDb.OleDbType.VarWChar, 0, "Description")});
            // 
            // oleDbUpdateCommand1
            // 
            this.oleDbUpdateCommand1.CommandText = resources.GetString("oleDbUpdateCommand1.CommandText");
            this.oleDbUpdateCommand1.Connection = this.dbconnection;
            this.oleDbUpdateCommand1.Parameters.AddRange(new System.Data.OleDb.OleDbParameter[] {
            new System.Data.OleDb.OleDbParameter("Caption", System.Data.OleDb.OleDbType.VarWChar, 255, "Caption"),
            new System.Data.OleDb.OleDbParameter("CategoryName", System.Data.OleDb.OleDbType.VarWChar, 255, "CategoryName"),
            new System.Data.OleDb.OleDbParameter("Description", System.Data.OleDb.OleDbType.VarWChar, 0, "Description"),
            new System.Data.OleDb.OleDbParameter("Original_CategoryId", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, false, ((byte)(0)), ((byte)(0)), "CategoryId", System.Data.DataRowVersion.Original, null),
            new System.Data.OleDb.OleDbParameter("Original_Caption", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, false, ((byte)(0)), ((byte)(0)), "Caption", System.Data.DataRowVersion.Original, null),
            new System.Data.OleDb.OleDbParameter("Original_Caption1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, false, ((byte)(0)), ((byte)(0)), "Caption", System.Data.DataRowVersion.Original, null),
            new System.Data.OleDb.OleDbParameter("Original_CategoryName", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, false, ((byte)(0)), ((byte)(0)), "CategoryName", System.Data.DataRowVersion.Original, null),
            new System.Data.OleDb.OleDbParameter("Original_CategoryName1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, false, ((byte)(0)), ((byte)(0)), "CategoryName", System.Data.DataRowVersion.Original, null)});
            // 
            // oleDbDeleteCommand1
            // 
            this.oleDbDeleteCommand1.CommandText = "DELETE FROM Categories WHERE (CategoryId = ?) AND (Caption = ? OR ? IS NULL AND C" +
    "aption IS NULL) AND (CategoryName = ? OR ? IS NULL AND CategoryName IS NULL)";
            this.oleDbDeleteCommand1.Connection = this.dbconnection;
            this.oleDbDeleteCommand1.Parameters.AddRange(new System.Data.OleDb.OleDbParameter[] {
            new System.Data.OleDb.OleDbParameter("Original_CategoryId", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, false, ((byte)(0)), ((byte)(0)), "CategoryId", System.Data.DataRowVersion.Original, null),
            new System.Data.OleDb.OleDbParameter("Original_Caption", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, false, ((byte)(0)), ((byte)(0)), "Caption", System.Data.DataRowVersion.Original, null),
            new System.Data.OleDb.OleDbParameter("Original_Caption1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, false, ((byte)(0)), ((byte)(0)), "Caption", System.Data.DataRowVersion.Original, null),
            new System.Data.OleDb.OleDbParameter("Original_CategoryName", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, false, ((byte)(0)), ((byte)(0)), "CategoryName", System.Data.DataRowVersion.Original, null),
            new System.Data.OleDb.OleDbParameter("Original_CategoryName1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, false, ((byte)(0)), ((byte)(0)), "CategoryName", System.Data.DataRowVersion.Original, null)});
            // 
            // categoriesAdapter
            // 
            this.categoriesAdapter.DeleteCommand = this.oleDbDeleteCommand1;
            this.categoriesAdapter.InsertCommand = this.oleDbInsertCommand1;
            this.categoriesAdapter.SelectCommand = this.oleDbSelectCommand1;
            this.categoriesAdapter.TableMappings.AddRange(new System.Data.Common.DataTableMapping[] {
            new System.Data.Common.DataTableMapping("Table", "Categories", new System.Data.Common.DataColumnMapping[] {
                        new System.Data.Common.DataColumnMapping("Caption", "Caption"),
                        new System.Data.Common.DataColumnMapping("CategoryId", "CategoryId"),
                        new System.Data.Common.DataColumnMapping("CategoryName", "CategoryName"),
                        new System.Data.Common.DataColumnMapping("Description", "Description")})});
            this.categoriesAdapter.UpdateCommand = this.oleDbUpdateCommand1;
            // 
            // featuresAdapter
            // 
            this.featuresAdapter.DeleteCommand = this.oleDbDeleteCommand2;
            this.featuresAdapter.InsertCommand = this.oleDbInsertCommand2;
            this.featuresAdapter.SelectCommand = this.oleDbSelectCommand2;
            this.featuresAdapter.TableMappings.AddRange(new System.Data.Common.DataTableMapping[] {
            new System.Data.Common.DataTableMapping("Table", "Features", new System.Data.Common.DataColumnMapping[] {
                        new System.Data.Common.DataColumnMapping("Caption", "Caption"),
                        new System.Data.Common.DataColumnMapping("CategoryId", "CategoryId"),
                        new System.Data.Common.DataColumnMapping("Description", "Description"),
                        new System.Data.Common.DataColumnMapping("FeaturesId", "FeaturesId")})});
            this.featuresAdapter.UpdateCommand = this.oleDbUpdateCommand2;
            // 
            // oleDbDeleteCommand2
            // 
            this.oleDbDeleteCommand2.CommandText = "DELETE FROM Features WHERE (FeaturesId = ?) AND (Caption = ?) AND (CategoryId = ?" +
    ")";
            this.oleDbDeleteCommand2.Parameters.AddRange(new System.Data.OleDb.OleDbParameter[] {
            new System.Data.OleDb.OleDbParameter("Original_FeaturesId", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, false, ((byte)(0)), ((byte)(0)), "FeaturesId", System.Data.DataRowVersion.Original, null),
            new System.Data.OleDb.OleDbParameter("Original_Caption", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, false, ((byte)(0)), ((byte)(0)), "Caption", System.Data.DataRowVersion.Original, null),
            new System.Data.OleDb.OleDbParameter("Original_CategoryId", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, false, ((byte)(0)), ((byte)(0)), "CategoryId", System.Data.DataRowVersion.Original, null)});
            // 
            // oleDbInsertCommand2
            // 
            this.oleDbInsertCommand2.CommandText = "INSERT INTO Features(Caption, CategoryId, Description) VALUES (?, ?, ?)";
            this.oleDbInsertCommand2.Parameters.AddRange(new System.Data.OleDb.OleDbParameter[] {
            new System.Data.OleDb.OleDbParameter("Caption", System.Data.OleDb.OleDbType.VarWChar, 255, "Caption"),
            new System.Data.OleDb.OleDbParameter("CategoryId", System.Data.OleDb.OleDbType.Integer, 0, "CategoryId"),
            new System.Data.OleDb.OleDbParameter("Description", System.Data.OleDb.OleDbType.VarWChar, 0, "Description")});
            // 
            // oleDbSelectCommand2
            // 
            this.oleDbSelectCommand2.CommandText = "SELECT Caption, CategoryId, Description, FeaturesId FROM Features order by Positi" +
    "onInSheet";
            this.oleDbSelectCommand2.Connection = this.dbconnection;
            // 
            // oleDbUpdateCommand2
            // 
            this.oleDbUpdateCommand2.CommandText = "UPDATE Features SET Caption = ?, CategoryId = ?, Description = ? WHERE (FeaturesI" +
    "d = ?) AND (Caption = ?) AND (CategoryId = ?)";
            this.oleDbUpdateCommand2.Parameters.AddRange(new System.Data.OleDb.OleDbParameter[] {
            new System.Data.OleDb.OleDbParameter("Caption", System.Data.OleDb.OleDbType.VarWChar, 255, "Caption"),
            new System.Data.OleDb.OleDbParameter("CategoryId", System.Data.OleDb.OleDbType.Integer, 0, "CategoryId"),
            new System.Data.OleDb.OleDbParameter("Description", System.Data.OleDb.OleDbType.VarWChar, 0, "Description"),
            new System.Data.OleDb.OleDbParameter("Original_FeaturesId", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, false, ((byte)(0)), ((byte)(0)), "FeaturesId", System.Data.DataRowVersion.Original, null),
            new System.Data.OleDb.OleDbParameter("Original_Caption", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, false, ((byte)(0)), ((byte)(0)), "Caption", System.Data.DataRowVersion.Original, null),
            new System.Data.OleDb.OleDbParameter("Original_CategoryId", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, false, ((byte)(0)), ((byte)(0)), "CategoryId", System.Data.DataRowVersion.Original, null)});
            // 
            // oleDbSelectCommand3
            // 
            this.oleDbSelectCommand3.CommandText = "SELECT FeaturesId, HiperlinksId, LinkName, Url FROM Hyperlinks";
            this.oleDbSelectCommand3.Connection = this.dbconnection;
            // 
            // hyperlinksAdapter
            // 
            this.hyperlinksAdapter.SelectCommand = this.oleDbSelectCommand3;
            this.hyperlinksAdapter.TableMappings.AddRange(new System.Data.Common.DataTableMapping[] {
            new System.Data.Common.DataTableMapping("Table", "Hyperlinks", new System.Data.Common.DataColumnMapping[] {
                        new System.Data.Common.DataColumnMapping("FeaturesId", "FeaturesId"),
                        new System.Data.Common.DataColumnMapping("HiperlinksId", "HiperlinksId"),
                        new System.Data.Common.DataColumnMapping("LinkName", "LinkName"),
                        new System.Data.Common.DataColumnMapping("Url", "Url")})});
            // 
            // mainToolbar
            // 
            this.mainToolbar.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripButton4,
            this.toolStripButton3,
            this.toolStripButton2,
            this.toolStripButton1});
            this.mainToolbar.Location = new System.Drawing.Point(0, 0);
            this.mainToolbar.Name = "mainToolbar";
            this.mainToolbar.Size = new System.Drawing.Size(528, 38);
            this.mainToolbar.TabIndex = 5;
            this.mainToolbar.Text = "mainToolbar";
            // 
            // toolStripButton4
            // 
            this.toolStripButton4.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton4.Image")));
            this.toolStripButton4.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton4.Name = "toolStripButton4";
            this.toolStripButton4.Size = new System.Drawing.Size(78, 35);
            this.toolStripButton4.Text = "Save to Excel";
            this.toolStripButton4.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.toolStripButton4.Click += new System.EventHandler(this.btnExportExcel_Click);
            // 
            // toolStripButton3
            // 
            this.toolStripButton3.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton3.Image")));
            this.toolStripButton3.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton3.Name = "toolStripButton3";
            this.toolStripButton3.Size = new System.Drawing.Size(94, 35);
            this.toolStripButton3.Text = "Export to HTML";
            this.toolStripButton3.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.toolStripButton3.Click += new System.EventHandler(this.btnExportHtml_Click);
            // 
            // toolStripButton2
            // 
            this.toolStripButton2.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton2.Image")));
            this.toolStripButton2.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton2.Name = "toolStripButton2";
            this.toolStripButton2.Size = new System.Drawing.Size(82, 35);
            this.toolStripButton2.Text = "Export to PDF";
            this.toolStripButton2.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.toolStripButton2.Click += new System.EventHandler(this.btnExportPDF_Click);
            // 
            // toolStripButton1
            // 
            this.toolStripButton1.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.toolStripButton1.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton1.Image")));
            this.toolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton1.Name = "toolStripButton1";
            this.toolStripButton1.Size = new System.Drawing.Size(59, 35);
            this.toolStripButton1.Text = "     E&xit     ";
            this.toolStripButton1.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.toolStripButton1.Click += new System.EventHandler(this.button2_Click);
            // 
            // mainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(528, 126);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.mainToolbar);
            this.Name = "mainForm";
            this.Text = "Features FlexCel";
            this.mainToolbar.ResumeLayout(false);
            this.mainToolbar.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        private ToolStrip mainToolbar;
        private ToolStripButton toolStripButton4;
        private ToolStripButton toolStripButton3;
        private ToolStripButton toolStripButton2;
        private ToolStripButton toolStripButton1;
    }
}

