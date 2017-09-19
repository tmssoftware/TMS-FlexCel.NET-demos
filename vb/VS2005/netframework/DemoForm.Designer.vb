Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports System.Globalization
Imports System.Resources
Imports System.Threading
Namespace MainDemo
	Partial Public Class DemoForm
		Inherits System.Windows.Forms.Form

		Private panel1 As System.Windows.Forms.Panel
		Private panel3 As System.Windows.Forms.Panel
		Private splitter1 As System.Windows.Forms.Splitter
		Private statusBar1 As System.Windows.Forms.StatusBar
		Private mainMenu1 As System.Windows.Forms.MainMenu
		Private menuItem1 As System.Windows.Forms.MenuItem
		Private menuItem3 As System.Windows.Forms.MenuItem
		Private menuItem6 As System.Windows.Forms.MenuItem
		Private panel4 As System.Windows.Forms.Panel
		Private panel5 As System.Windows.Forms.Panel
		Private label1 As System.Windows.Forms.Label
		Private WithEvents descriptionText As System.Windows.Forms.RichTextBox
		Private WithEvents modulesList As System.Windows.Forms.TreeView
		Private pageSetupDialog1 As System.Windows.Forms.PageSetupDialog
		Private WithEvents menuRunSelected As System.Windows.Forms.MenuItem
		Private WithEvents menuAbout As System.Windows.Forms.MenuItem
		Private WithEvents menuViewTemplate As System.Windows.Forms.MenuItem
		Private WithEvents menuExit As System.Windows.Forms.MenuItem
		Private menuItem4 As System.Windows.Forms.MenuItem
		Private WithEvents menuOpenProject As System.Windows.Forms.MenuItem
		Private SearchTip As System.Windows.Forms.ToolTip
		Private components As System.ComponentModel.IContainer = Nothing

		''' <summary>
		''' Clean up any resources being used.
		''' </summary>
		Protected Overrides Sub Dispose(ByVal disposing As Boolean)
			If disposing Then
				If components IsNot Nothing Then
					components.Dispose()
				End If
			End If
			MyBase.Dispose(disposing)
		End Sub

		#Region "Windows Form Designer generated code"
		''' <summary>
		''' Required method for Designer support - do not modify
		''' the contents of this method with the code editor.
		''' </summary>
		Private Sub InitializeComponent()
			Me.components = New System.ComponentModel.Container()
			Dim resources As New System.ComponentModel.ComponentResourceManager(GetType(DemoForm))
			Me.panel1 = New System.Windows.Forms.Panel()
			Me.modulesList = New System.Windows.Forms.TreeView()
			Me.panel5 = New System.Windows.Forms.Panel()
			Me.label1 = New System.Windows.Forms.Label()
			Me.panel3 = New System.Windows.Forms.Panel()
			Me.descriptionText = New System.Windows.Forms.RichTextBox()
			Me.splitter1 = New System.Windows.Forms.Splitter()
			Me.panel4 = New System.Windows.Forms.Panel()
			Me.statusBar1 = New System.Windows.Forms.StatusBar()
			Me.mainMenu1 = New System.Windows.Forms.MainMenu(Me.components)
			Me.menuItem1 = New System.Windows.Forms.MenuItem()
			Me.menuExit = New System.Windows.Forms.MenuItem()
			Me.menuItem3 = New System.Windows.Forms.MenuItem()
			Me.menuRunSelected = New System.Windows.Forms.MenuItem()
			Me.menuItem4 = New System.Windows.Forms.MenuItem()
			Me.menuViewTemplate = New System.Windows.Forms.MenuItem()
			Me.menuOpenProject = New System.Windows.Forms.MenuItem()
			Me.menuItem6 = New System.Windows.Forms.MenuItem()
			Me.menuAbout = New System.Windows.Forms.MenuItem()
			Me.pageSetupDialog1 = New System.Windows.Forms.PageSetupDialog()
			Me.SearchTip = New System.Windows.Forms.ToolTip(Me.components)
			Me.mainToolbar = New System.Windows.Forms.ToolStrip()
			Me.toolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
			Me.toolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
			Me.sdSearch = New System.Windows.Forms.ToolStripTextBox()
			Me.btnRunSelected = New System.Windows.Forms.ToolStripButton()
			Me.btnViewTemplate = New System.Windows.Forms.ToolStripButton()
			Me.btnOpenFolder = New System.Windows.Forms.ToolStripButton()
			Me.btnOpenProject = New System.Windows.Forms.ToolStripButton()
			Me.btnExit = New System.Windows.Forms.ToolStripButton()
			Me.btnAbout = New System.Windows.Forms.ToolStripButton()
			Me.panel1.SuspendLayout()
			Me.panel5.SuspendLayout()
			Me.panel3.SuspendLayout()
			Me.mainToolbar.SuspendLayout()
			Me.SuspendLayout()
			' 
			' panel1
			' 
			Me.panel1.Controls.Add(Me.modulesList)
			Me.panel1.Controls.Add(Me.panel5)
			Me.panel1.Dock = System.Windows.Forms.DockStyle.Left
			Me.panel1.Location = New System.Drawing.Point(0, 46)
			Me.panel1.Name = "panel1"
			Me.panel1.Size = New System.Drawing.Size(232, 446)
			Me.panel1.TabIndex = 0
			' 
			' modulesList
			' 
			Me.modulesList.Dock = System.Windows.Forms.DockStyle.Fill
			Me.modulesList.HideSelection = False
			Me.modulesList.Location = New System.Drawing.Point(0, 24)
			Me.modulesList.Name = "modulesList"
			Me.modulesList.Size = New System.Drawing.Size(232, 422)
			Me.modulesList.TabIndex = 4
'			Me.modulesList.AfterSelect += New System.Windows.Forms.TreeViewEventHandler(Me.modulesList_AfterSelect)
			' 
			' panel5
			' 
			Me.panel5.BackColor = System.Drawing.Color.FromArgb((CInt((CByte(64)))), (CInt((CByte(64)))), (CInt((CByte(64)))))
			Me.panel5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
			Me.panel5.Controls.Add(Me.label1)
			Me.panel5.Dock = System.Windows.Forms.DockStyle.Top
			Me.panel5.ForeColor = System.Drawing.Color.White
			Me.panel5.Location = New System.Drawing.Point(0, 0)
			Me.panel5.Name = "panel5"
			Me.panel5.Size = New System.Drawing.Size(232, 24)
			Me.panel5.TabIndex = 3
			' 
			' label1
			' 
			Me.label1.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label1.Location = New System.Drawing.Point(1, 4)
			Me.label1.Name = "label1"
			Me.label1.Size = New System.Drawing.Size(100, 23)
			Me.label1.TabIndex = 0
			Me.label1.Text = "Select a Demo"
			' 
			' panel3
			' 
			Me.panel3.Controls.Add(Me.descriptionText)
			Me.panel3.Controls.Add(Me.splitter1)
			Me.panel3.Controls.Add(Me.panel4)
			Me.panel3.Dock = System.Windows.Forms.DockStyle.Fill
			Me.panel3.Location = New System.Drawing.Point(232, 46)
			Me.panel3.Name = "panel3"
			Me.panel3.Size = New System.Drawing.Size(616, 446)
			Me.panel3.TabIndex = 2
			' 
			' descriptionText
			' 
			Me.descriptionText.BackColor = System.Drawing.SystemColors.Window
			Me.descriptionText.Dock = System.Windows.Forms.DockStyle.Fill
			Me.descriptionText.Location = New System.Drawing.Point(3, 24)
			Me.descriptionText.Name = "descriptionText"
			Me.descriptionText.ReadOnly = True
			Me.descriptionText.Size = New System.Drawing.Size(613, 422)
			Me.descriptionText.TabIndex = 1
			Me.descriptionText.Text = ""
'			Me.descriptionText.LinkClicked += New System.Windows.Forms.LinkClickedEventHandler(Me.descriptionText_LinkClicked)
			' 
			' splitter1
			' 
			Me.splitter1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.splitter1.Location = New System.Drawing.Point(0, 24)
			Me.splitter1.Name = "splitter1"
			Me.splitter1.Size = New System.Drawing.Size(3, 422)
			Me.splitter1.TabIndex = 0
			Me.splitter1.TabStop = False
			' 
			' panel4
			' 
			Me.panel4.BackColor = System.Drawing.Color.FromArgb((CInt((CByte(64)))), (CInt((CByte(64)))), (CInt((CByte(64)))))
			Me.panel4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
			Me.panel4.Dock = System.Windows.Forms.DockStyle.Top
			Me.panel4.ForeColor = System.Drawing.Color.White
			Me.panel4.Location = New System.Drawing.Point(0, 0)
			Me.panel4.Name = "panel4"
			Me.panel4.Size = New System.Drawing.Size(616, 24)
			Me.panel4.TabIndex = 2
			Me.panel4.Visible = False
			' 
			' statusBar1
			' 
			Me.statusBar1.Location = New System.Drawing.Point(0, 492)
			Me.statusBar1.Name = "statusBar1"
			Me.statusBar1.Size = New System.Drawing.Size(848, 22)
			Me.statusBar1.TabIndex = 3
			Me.statusBar1.Text = "statusBar1"
			' 
			' mainMenu1
			' 
			Me.mainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() { Me.menuItem1, Me.menuItem3, Me.menuItem6})
			' 
			' menuItem1
			' 
			Me.menuItem1.Index = 0
			Me.menuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() { Me.menuExit})
			Me.menuItem1.Text = "File"
			' 
			' menuExit
			' 
			Me.menuExit.Index = 0
			Me.menuExit.Text = "Exit"
'			Me.menuExit.Click += New System.EventHandler(Me.Exit_Click)
			' 
			' menuItem3
			' 
			Me.menuItem3.Index = 1
			Me.menuItem3.MenuItems.AddRange(New System.Windows.Forms.MenuItem() { Me.menuRunSelected, Me.menuItem4, Me.menuViewTemplate, Me.menuOpenProject})
			Me.menuItem3.Text = "Demo"
			' 
			' menuRunSelected
			' 
			Me.menuRunSelected.Index = 0
			Me.menuRunSelected.Shortcut = System.Windows.Forms.Shortcut.F5
			Me.menuRunSelected.Text = "Run Selected"
'			Me.menuRunSelected.Click += New System.EventHandler(Me.RunSelected_Click)
			' 
			' menuItem4
			' 
			Me.menuItem4.Index = 1
			Me.menuItem4.Text = "-"
			' 
			' menuViewTemplate
			' 
			Me.menuViewTemplate.Index = 2
			Me.menuViewTemplate.Text = "View Template"
'			Me.menuViewTemplate.Click += New System.EventHandler(Me.ViewTemplate_Click)
			' 
			' menuOpenProject
			' 
			Me.menuOpenProject.Index = 3
			Me.menuOpenProject.Text = "Open Project"
'			Me.menuOpenProject.Click += New System.EventHandler(Me.btnOpenProject_Click)
			' 
			' menuItem6
			' 
			Me.menuItem6.Index = 2
			Me.menuItem6.MenuItems.AddRange(New System.Windows.Forms.MenuItem() { Me.menuAbout})
			Me.menuItem6.Text = "Help"
			' 
			' menuAbout
			' 
			Me.menuAbout.Index = 0
			Me.menuAbout.Text = "About..."
'			Me.menuAbout.Click += New System.EventHandler(Me.About_Click)
			' 
			' mainToolbar
			' 
			Me.mainToolbar.ImageScalingSize = New System.Drawing.Size(24, 24)
			Me.mainToolbar.Items.AddRange(New System.Windows.Forms.ToolStripItem() { Me.btnRunSelected, Me.toolStripSeparator2, Me.btnViewTemplate, Me.btnOpenFolder, Me.btnOpenProject, Me.toolStripSeparator1, Me.btnExit, Me.btnAbout, Me.sdSearch})
			Me.mainToolbar.Location = New System.Drawing.Point(0, 0)
			Me.mainToolbar.Name = "mainToolbar"
			Me.mainToolbar.Size = New System.Drawing.Size(848, 46)
			Me.mainToolbar.TabIndex = 11
			Me.mainToolbar.Text = "toolStrip1"
			' 
			' toolStripSeparator2
			' 
			Me.toolStripSeparator2.Name = "toolStripSeparator2"
			Me.toolStripSeparator2.Size = New System.Drawing.Size(6, 46)
			' 
			' toolStripSeparator1
			' 
			Me.toolStripSeparator1.Name = "toolStripSeparator1"
			Me.toolStripSeparator1.Size = New System.Drawing.Size(6, 46)
			' 
			' sdSearch
			' 
			Me.sdSearch.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
			Me.sdSearch.Margin = New System.Windows.Forms.Padding(1, 0, 20, 0)
			Me.sdSearch.Name = "sdSearch"
			Me.sdSearch.Size = New System.Drawing.Size(160, 46)
'			Me.sdSearch.Enter += New System.EventHandler(Me.sdSearch_Enter)
'			Me.sdSearch.Leave += New System.EventHandler(Me.sdSearch_Leave)
'			Me.sdSearch.TextChanged += New System.EventHandler(Me.sdSearch_TextChanged)
			' 
			' btnRunSelected
			' 
			Me.btnRunSelected.Image = (CType(resources.GetObject("btnRunSelected.Image"), System.Drawing.Image))
			Me.btnRunSelected.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnRunSelected.Name = "btnRunSelected"
			Me.btnRunSelected.Size = New System.Drawing.Size(79, 43)
			Me.btnRunSelected.Text = "&Run Selected"
			Me.btnRunSelected.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
'			Me.btnRunSelected.Click += New System.EventHandler(Me.RunSelected_Click)
			' 
			' btnViewTemplate
			' 
			Me.btnViewTemplate.Image = (CType(resources.GetObject("btnViewTemplate.Image"), System.Drawing.Image))
			Me.btnViewTemplate.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnViewTemplate.Name = "btnViewTemplate"
			Me.btnViewTemplate.Size = New System.Drawing.Size(89, 43)
			Me.btnViewTemplate.Text = "View &Template"
			Me.btnViewTemplate.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
'			Me.btnViewTemplate.Click += New System.EventHandler(Me.ViewTemplate_Click)
			' 
			' btnOpenFolder
			' 
			Me.btnOpenFolder.Image = (CType(resources.GetObject("btnOpenFolder.Image"), System.Drawing.Image))
			Me.btnOpenFolder.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnOpenFolder.Name = "btnOpenFolder"
			Me.btnOpenFolder.Size = New System.Drawing.Size(76, 43)
			Me.btnOpenFolder.Text = "&Open Folder"
			Me.btnOpenFolder.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
'			Me.btnOpenFolder.Click += New System.EventHandler(Me.btnOpenFolder_Click)
			' 
			' btnOpenProject
			' 
			Me.btnOpenProject.Image = (CType(resources.GetObject("btnOpenProject.Image"), System.Drawing.Image))
			Me.btnOpenProject.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnOpenProject.Name = "btnOpenProject"
			Me.btnOpenProject.Size = New System.Drawing.Size(80, 43)
			Me.btnOpenProject.Text = "Open &Project"
			Me.btnOpenProject.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
'			Me.btnOpenProject.Click += New System.EventHandler(Me.btnOpenProject_Click)
			' 
			' btnExit
			' 
			Me.btnExit.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
			Me.btnExit.Image = (CType(resources.GetObject("btnExit.Image"), System.Drawing.Image))
			Me.btnExit.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnExit.Name = "btnExit"
			Me.btnExit.Size = New System.Drawing.Size(59, 43)
			Me.btnExit.Text = "     E&xit     "
			Me.btnExit.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
'			Me.btnExit.Click += New System.EventHandler(Me.Exit_Click)
			' 
			' btnAbout
			' 
			Me.btnAbout.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
			Me.btnAbout.Image = (CType(resources.GetObject("btnAbout.Image"), System.Drawing.Image))
			Me.btnAbout.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnAbout.Name = "btnAbout"
			Me.btnAbout.Size = New System.Drawing.Size(56, 43)
			Me.btnAbout.Text = "  About  "
			Me.btnAbout.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
'			Me.btnAbout.Click += New System.EventHandler(Me.About_Click)
			' 
			' DemoForm
			' 
			Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New System.Drawing.Size(848, 514)
			Me.Controls.Add(Me.panel3)
			Me.Controls.Add(Me.panel1)
			Me.Controls.Add(Me.statusBar1)
			Me.Controls.Add(Me.mainToolbar)
			Me.Icon = (CType(resources.GetObject("$this.Icon"), System.Drawing.Icon))
			Me.Menu = Me.mainMenu1
			Me.Name = "DemoForm"
			Me.Text = "FlexCel Well"
			Me.panel1.ResumeLayout(False)
			Me.panel5.ResumeLayout(False)
			Me.panel3.ResumeLayout(False)
			Me.mainToolbar.ResumeLayout(False)
			Me.mainToolbar.PerformLayout()
			Me.ResumeLayout(False)
			Me.PerformLayout()

		End Sub
		#End Region

		Private mainToolbar As ToolStrip
		Private WithEvents btnRunSelected As ToolStripButton
		Private toolStripSeparator2 As ToolStripSeparator
		Private WithEvents btnViewTemplate As ToolStripButton
		Private WithEvents btnOpenFolder As ToolStripButton
		Private WithEvents btnOpenProject As ToolStripButton
		Private toolStripSeparator1 As ToolStripSeparator
		Private WithEvents btnAbout As ToolStripButton
		Private WithEvents btnExit As ToolStripButton
		Private WithEvents sdSearch As ToolStripTextBox
	End Class
End Namespace


