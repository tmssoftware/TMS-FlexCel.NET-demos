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
using System.Threading;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using FlexCel.Report;

namespace GenericReports2
{
    public partial class mainForm: System.Windows.Forms.Form
    {
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Data.DataSet dataSet;
        private System.Windows.Forms.DataGrid dataGrid;
        private System.Data.OleDb.OleDbConnection Connection;
        private System.Data.OleDb.OleDbDataAdapter dbDataAdapter;
        private FlexCel.Report.FlexCelReport Report;
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
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.Report = new FlexCel.Report.FlexCelReport();
            this.Connection = new System.Data.OleDb.OleDbConnection();
            this.dataSet = new System.Data.DataSet();
            this.dataGrid = new System.Windows.Forms.DataGrid();
            this.dbDataAdapter = new System.Data.OleDb.OleDbDataAdapter();
            this.mainToolbar = new System.Windows.Forms.ToolStrip();
            this.btnOpenConnection = new System.Windows.Forms.ToolStripButton();
            this.btnQuery = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.btnExportExcel = new System.Windows.Forms.ToolStripButton();
            this.btnExit = new System.Windows.Forms.ToolStripButton();
            ((System.ComponentModel.ISupportInitialize)(this.dataSet)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid)).BeginInit();
            this.mainToolbar.SuspendLayout();
            this.SuspendLayout();
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|Excel 97/2003|*.xls|Excel 2007|*.xlsx;*.xlsm|All " +
    "files|*.*";
            this.saveFileDialog1.RestoreDirectory = true;
            // 
            // Report
            // 
            this.Report.AllowOverwritingFiles = true;
            this.Report.DeleteEmptyBands = FlexCel.Report.TDeleteEmptyBands.ClearDataAndFormats;
            this.Report.DeleteEmptyRanges = false;
            // 
            // dataSet
            // 
            this.dataSet.DataSetName = "NewDataSet";
            this.dataSet.Locale = new System.Globalization.CultureInfo("es-ES");
            // 
            // dataGrid
            // 
            this.dataGrid.DataMember = "";
            this.dataGrid.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGrid.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.dataGrid.Location = new System.Drawing.Point(0, 38);
            this.dataGrid.Name = "dataGrid";
            this.dataGrid.Size = new System.Drawing.Size(528, 239);
            this.dataGrid.TabIndex = 4;
            // 
            // mainToolbar
            // 
            this.mainToolbar.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnOpenConnection,
            this.btnQuery,
            this.toolStripSeparator1,
            this.btnExportExcel,
            this.btnExit});
            this.mainToolbar.Location = new System.Drawing.Point(0, 0);
            this.mainToolbar.Name = "mainToolbar";
            this.mainToolbar.Size = new System.Drawing.Size(528, 38);
            this.mainToolbar.TabIndex = 11;
            this.mainToolbar.Text = "mainToolbar";
            // 
            // btnOpenConnection
            // 
            this.btnOpenConnection.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenConnection.Image")));
            this.btnOpenConnection.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnOpenConnection.Name = "btnOpenConnection";
            this.btnOpenConnection.Size = new System.Drawing.Size(103, 35);
            this.btnOpenConnection.Text = "Open connection";
            this.btnOpenConnection.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnOpenConnection.Click += new System.EventHandler(this.btnOpenconnection_Click);
            // 
            // btnQuery
            // 
            this.btnQuery.Image = ((System.Drawing.Image)(resources.GetObject("btnQuery.Image")));
            this.btnQuery.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnQuery.Name = "btnQuery";
            this.btnQuery.Size = new System.Drawing.Size(70, 35);
            this.btnQuery.Text = "Query Data";
            this.btnQuery.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnQuery.Click += new System.EventHandler(this.btnQuery_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 38);
            // 
            // btnExportExcel
            // 
            this.btnExportExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExportExcel.Image")));
            this.btnExportExcel.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnExportExcel.Name = "btnExportExcel";
            this.btnExportExcel.Size = new System.Drawing.Size(87, 35);
            this.btnExportExcel.Text = "Export to Excel";
            this.btnExportExcel.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnExportExcel.Click += new System.EventHandler(this.btnExportExcel_Click);
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
            // mainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(528, 277);
            this.Controls.Add(this.dataGrid);
            this.Controls.Add(this.mainToolbar);
            this.Name = "mainForm";
            this.Text = "Generic Reports 2";
            ((System.ComponentModel.ISupportInitialize)(this.dataSet)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid)).EndInit();
            this.mainToolbar.ResumeLayout(false);
            this.mainToolbar.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        private ToolStrip mainToolbar;
        private ToolStripButton btnOpenConnection;
        private ToolStripButton btnQuery;
        private ToolStripSeparator toolStripSeparator1;
        private ToolStripButton btnExportExcel;
        private ToolStripButton btnExit;
    }
}

