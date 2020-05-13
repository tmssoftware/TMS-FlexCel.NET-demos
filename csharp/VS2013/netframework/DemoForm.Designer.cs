using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Reflection;
using System.Globalization;
using System.Resources;
using System.Threading;
namespace MainDemo
{
    public partial class DemoForm: System.Windows.Forms.Form
    {
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Splitter splitter1;
        private System.Windows.Forms.StatusBar statusBar1;
        private System.Windows.Forms.MainMenu mainMenu1;
        private System.Windows.Forms.MenuItem menuItem1;
        private System.Windows.Forms.MenuItem menuItem3;
        private System.Windows.Forms.MenuItem menuItem6;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TreeView modulesList;
        private System.Windows.Forms.PageSetupDialog pageSetupDialog1;
        private System.Windows.Forms.MenuItem menuRunSelected;
        private System.Windows.Forms.MenuItem menuAbout;
        private System.Windows.Forms.MenuItem menuViewTemplate;
        private System.Windows.Forms.MenuItem menuExit;
        private System.Windows.Forms.MenuItem menuItem4;
        private System.Windows.Forms.MenuItem menuOpenProject;
        private System.Windows.Forms.ToolTip SearchTip;
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DemoForm));
            this.panel1 = new System.Windows.Forms.Panel();
            this.modulesList = new System.Windows.Forms.TreeView();
            this.panel5 = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.panel3 = new System.Windows.Forms.Panel();
            this.splitter1 = new System.Windows.Forms.Splitter();
            this.panel4 = new System.Windows.Forms.Panel();
            this.statusBar1 = new System.Windows.Forms.StatusBar();
            this.mainMenu1 = new System.Windows.Forms.MainMenu(this.components);
            this.menuItem1 = new System.Windows.Forms.MenuItem();
            this.menuExit = new System.Windows.Forms.MenuItem();
            this.menuItem3 = new System.Windows.Forms.MenuItem();
            this.menuRunSelected = new System.Windows.Forms.MenuItem();
            this.menuItem4 = new System.Windows.Forms.MenuItem();
            this.menuViewTemplate = new System.Windows.Forms.MenuItem();
            this.menuOpenProject = new System.Windows.Forms.MenuItem();
            this.menuItem6 = new System.Windows.Forms.MenuItem();
            this.menuAbout = new System.Windows.Forms.MenuItem();
            this.pageSetupDialog1 = new System.Windows.Forms.PageSetupDialog();
            this.SearchTip = new System.Windows.Forms.ToolTip(this.components);
            this.mainToolbar = new System.Windows.Forms.ToolStrip();
            this.btnRunSelected = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.btnViewTemplate = new System.Windows.Forms.ToolStripButton();
            this.btnOpenFolder = new System.Windows.Forms.ToolStripButton();
            this.btnOpenProject = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.btnExit = new System.Windows.Forms.ToolStripButton();
            this.btnAbout = new System.Windows.Forms.ToolStripButton();
            this.sdSearch = new System.Windows.Forms.ToolStripTextBox();
            this.panel1.SuspendLayout();
            this.panel5.SuspendLayout();
            this.panel3.SuspendLayout();
            this.mainToolbar.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.modulesList);
            this.panel1.Controls.Add(this.panel5);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel1.Location = new System.Drawing.Point(0, 63);
            this.panel1.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(464, 883);
            this.panel1.TabIndex = 0;
            // 
            // modulesList
            // 
            this.modulesList.Dock = System.Windows.Forms.DockStyle.Fill;
            this.modulesList.HideSelection = false;
            this.modulesList.Location = new System.Drawing.Point(0, 42);
            this.modulesList.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.modulesList.Name = "modulesList";
            this.modulesList.Size = new System.Drawing.Size(464, 841);
            this.modulesList.TabIndex = 4;
            this.modulesList.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.modulesList_AfterSelect);
            // 
            // panel5
            // 
            this.panel5.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.panel5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel5.Controls.Add(this.label1);
            this.panel5.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel5.ForeColor = System.Drawing.Color.White;
            this.panel5.Location = new System.Drawing.Point(0, 0);
            this.panel5.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(464, 42);
            this.panel5.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(2, 8);
            this.label1.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(200, 44);
            this.label1.TabIndex = 0;
            this.label1.Text = "Select a Demo";
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.splitter1);
            this.panel3.Controls.Add(this.panel4);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel3.Location = new System.Drawing.Point(464, 63);
            this.panel3.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(1232, 883);
            this.panel3.TabIndex = 2;
            // 
            // splitter1
            // 
            this.splitter1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.splitter1.Location = new System.Drawing.Point(0, 42);
            this.splitter1.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.splitter1.Name = "splitter1";
            this.splitter1.Size = new System.Drawing.Size(4, 841);
            this.splitter1.TabIndex = 0;
            this.splitter1.TabStop = false;
            // 
            // panel4
            // 
            this.panel4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.panel4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel4.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel4.ForeColor = System.Drawing.Color.White;
            this.panel4.Location = new System.Drawing.Point(0, 0);
            this.panel4.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(1232, 42);
            this.panel4.TabIndex = 2;
            this.panel4.Visible = false;
            // 
            // statusBar1
            // 
            this.statusBar1.Location = new System.Drawing.Point(0, 946);
            this.statusBar1.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.statusBar1.Name = "statusBar1";
            this.statusBar1.Size = new System.Drawing.Size(1696, 42);
            this.statusBar1.TabIndex = 3;
            this.statusBar1.Text = "statusBar1";
            // 
            // mainMenu1
            // 
            this.mainMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.menuItem1,
            this.menuItem3,
            this.menuItem6});
            // 
            // menuItem1
            // 
            this.menuItem1.Index = 0;
            this.menuItem1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.menuExit});
            this.menuItem1.Text = "File";
            // 
            // menuExit
            // 
            this.menuExit.Index = 0;
            this.menuExit.Text = "Exit";
            this.menuExit.Click += new System.EventHandler(this.Exit_Click);
            // 
            // menuItem3
            // 
            this.menuItem3.Index = 1;
            this.menuItem3.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.menuRunSelected,
            this.menuItem4,
            this.menuViewTemplate,
            this.menuOpenProject});
            this.menuItem3.Text = "Demo";
            // 
            // menuRunSelected
            // 
            this.menuRunSelected.Index = 0;
            this.menuRunSelected.Shortcut = System.Windows.Forms.Shortcut.F5;
            this.menuRunSelected.Text = "Run Selected";
            this.menuRunSelected.Click += new System.EventHandler(this.RunSelected_Click);
            // 
            // menuItem4
            // 
            this.menuItem4.Index = 1;
            this.menuItem4.Text = "-";
            // 
            // menuViewTemplate
            // 
            this.menuViewTemplate.Index = 2;
            this.menuViewTemplate.Text = "View Template";
            this.menuViewTemplate.Click += new System.EventHandler(this.ViewTemplate_Click);
            // 
            // menuOpenProject
            // 
            this.menuOpenProject.Index = 3;
            this.menuOpenProject.Text = "Open Project";
            this.menuOpenProject.Click += new System.EventHandler(this.btnOpenProject_Click);
            // 
            // menuItem6
            // 
            this.menuItem6.Index = 2;
            this.menuItem6.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.menuAbout});
            this.menuItem6.Text = "Help";
            // 
            // menuAbout
            // 
            this.menuAbout.Index = 0;
            this.menuAbout.Text = "About...";
            this.menuAbout.Click += new System.EventHandler(this.About_Click);
            // 
            // mainToolbar
            // 
            this.mainToolbar.ImageScalingSize = new System.Drawing.Size(24, 24);
            this.mainToolbar.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnRunSelected,
            this.toolStripSeparator2,
            this.btnViewTemplate,
            this.btnOpenFolder,
            this.btnOpenProject,
            this.toolStripSeparator1,
            this.btnExit,
            this.btnAbout,
            this.sdSearch});
            this.mainToolbar.Location = new System.Drawing.Point(0, 0);
            this.mainToolbar.Name = "mainToolbar";
            this.mainToolbar.Padding = new System.Windows.Forms.Padding(0, 0, 2, 0);
            this.mainToolbar.Size = new System.Drawing.Size(1696, 63);
            this.mainToolbar.TabIndex = 11;
            this.mainToolbar.Text = "toolStrip1";
            // 
            // btnRunSelected
            // 
            this.btnRunSelected.Image = ((System.Drawing.Image)(resources.GetObject("btnRunSelected.Image")));
            this.btnRunSelected.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnRunSelected.Name = "btnRunSelected";
            this.btnRunSelected.Size = new System.Drawing.Size(159, 60);
            this.btnRunSelected.Text = "&Run Selected";
            this.btnRunSelected.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnRunSelected.Click += new System.EventHandler(this.RunSelected_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(6, 63);
            // 
            // btnViewTemplate
            // 
            this.btnViewTemplate.Image = ((System.Drawing.Image)(resources.GetObject("btnViewTemplate.Image")));
            this.btnViewTemplate.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnViewTemplate.Name = "btnViewTemplate";
            this.btnViewTemplate.Size = new System.Drawing.Size(175, 60);
            this.btnViewTemplate.Text = "View &Template";
            this.btnViewTemplate.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnViewTemplate.Click += new System.EventHandler(this.ViewTemplate_Click);
            // 
            // btnOpenFolder
            // 
            this.btnOpenFolder.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenFolder.Image")));
            this.btnOpenFolder.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnOpenFolder.Name = "btnOpenFolder";
            this.btnOpenFolder.Size = new System.Drawing.Size(152, 60);
            this.btnOpenFolder.Text = "&Open Folder";
            this.btnOpenFolder.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnOpenFolder.Click += new System.EventHandler(this.btnOpenFolder_Click);
            // 
            // btnOpenProject
            // 
            this.btnOpenProject.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenProject.Image")));
            this.btnOpenProject.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnOpenProject.Name = "btnOpenProject";
            this.btnOpenProject.Size = new System.Drawing.Size(158, 60);
            this.btnOpenProject.Text = "Open &Project";
            this.btnOpenProject.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnOpenProject.Click += new System.EventHandler(this.btnOpenProject_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 63);
            // 
            // btnExit
            // 
            this.btnExit.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.btnExit.Image = ((System.Drawing.Image)(resources.GetObject("btnExit.Image")));
            this.btnExit.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(126, 60);
            this.btnExit.Text = "     E&xit     ";
            this.btnExit.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnExit.Click += new System.EventHandler(this.Exit_Click);
            // 
            // btnAbout
            // 
            this.btnAbout.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.btnAbout.Image = ((System.Drawing.Image)(resources.GetObject("btnAbout.Image")));
            this.btnAbout.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnAbout.Name = "btnAbout";
            this.btnAbout.Size = new System.Drawing.Size(112, 60);
            this.btnAbout.Text = "  About  ";
            this.btnAbout.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnAbout.Click += new System.EventHandler(this.About_Click);
            // 
            // sdSearch
            // 
            this.sdSearch.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.sdSearch.Margin = new System.Windows.Forms.Padding(1, 0, 20, 0);
            this.sdSearch.Name = "sdSearch";
            this.sdSearch.Size = new System.Drawing.Size(160, 63);
            this.sdSearch.Enter += new System.EventHandler(this.sdSearch_Enter);
            this.sdSearch.Leave += new System.EventHandler(this.sdSearch_Leave);
            this.sdSearch.TextChanged += new System.EventHandler(this.sdSearch_TextChanged);
            // 
            // DemoForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1696, 988);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.statusBar1);
            this.Controls.Add(this.mainToolbar);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.Menu = this.mainMenu1;
            this.Name = "DemoForm";
            this.Text = "FlexCel Well";
            this.panel1.ResumeLayout(false);
            this.panel5.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.mainToolbar.ResumeLayout(false);
            this.mainToolbar.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        private ToolStrip mainToolbar;
        private ToolStripButton btnRunSelected;
        private ToolStripSeparator toolStripSeparator2;
        private ToolStripButton btnViewTemplate;
        private ToolStripButton btnOpenFolder;
        private ToolStripButton btnOpenProject;
        private ToolStripSeparator toolStripSeparator1;
        private ToolStripButton btnAbout;
        private ToolStripButton btnExit;
        private ToolStripTextBox sdSearch;
    }
}


