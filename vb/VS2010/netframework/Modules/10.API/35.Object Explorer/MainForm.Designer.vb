Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Namespace ObjectExplorer
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Private openFileDialog1 As System.Windows.Forms.OpenFileDialog
		Private panel1 As System.Windows.Forms.Panel
		Private splitter1 As System.Windows.Forms.Splitter
		Private panel3 As System.Windows.Forms.Panel
		Private dataGrid As System.Windows.Forms.DataGrid
		Private panel4 As System.Windows.Forms.Panel
		Private splitter2 As System.Windows.Forms.Splitter
		Private PreviewBox As System.Windows.Forms.PictureBox
		Private saveImageDialog As System.Windows.Forms.SaveFileDialog
		Private lblObjects As System.Windows.Forms.Label
		Private WithEvents ObjTree As System.Windows.Forms.TreeView
		Private panel5 As System.Windows.Forms.Panel
		Private lblObjName As System.Windows.Forms.Label
		Private lblObjText As System.Windows.Forms.Label
		Private panel6 As System.Windows.Forms.Panel
		Private label1 As System.Windows.Forms.Label
		Private WithEvents cbSheet As System.Windows.Forms.ComboBox
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
			Dim resources As New System.ComponentModel.ComponentResourceManager(GetType(mainForm))
			Me.openFileDialog1 = New System.Windows.Forms.OpenFileDialog()
			Me.panel1 = New System.Windows.Forms.Panel()
			Me.ObjTree = New System.Windows.Forms.TreeView()
			Me.panel6 = New System.Windows.Forms.Panel()
			Me.cbSheet = New System.Windows.Forms.ComboBox()
			Me.label1 = New System.Windows.Forms.Label()
			Me.panel3 = New System.Windows.Forms.Panel()
			Me.lblObjects = New System.Windows.Forms.Label()
			Me.splitter1 = New System.Windows.Forms.Splitter()
			Me.dataGrid = New System.Windows.Forms.DataGrid()
			Me.panel4 = New System.Windows.Forms.Panel()
			Me.PreviewBox = New System.Windows.Forms.PictureBox()
			Me.splitter2 = New System.Windows.Forms.Splitter()
			Me.saveImageDialog = New System.Windows.Forms.SaveFileDialog()
			Me.panel5 = New System.Windows.Forms.Panel()
			Me.lblObjText = New System.Windows.Forms.Label()
			Me.lblObjName = New System.Windows.Forms.Label()
			Me.mainToolbar = New System.Windows.Forms.ToolStrip()
			Me.btnOpenFile = New System.Windows.Forms.ToolStripButton()
			Me.toolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
			Me.btnShowInExcel = New System.Windows.Forms.ToolStripButton()
			Me.btnSaveAsImage = New System.Windows.Forms.ToolStripButton()
			Me.btnExit = New System.Windows.Forms.ToolStripButton()
			Me.btnInfo = New System.Windows.Forms.ToolStripButton()
			Me.toolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
			Me.btnStretchPreview = New System.Windows.Forms.ToolStripButton()
			Me.panel1.SuspendLayout()
			Me.panel6.SuspendLayout()
			Me.panel3.SuspendLayout()
			CType(Me.dataGrid, System.ComponentModel.ISupportInitialize).BeginInit()
			Me.panel4.SuspendLayout()
			CType(Me.PreviewBox, System.ComponentModel.ISupportInitialize).BeginInit()
			Me.panel5.SuspendLayout()
			Me.mainToolbar.SuspendLayout()
			Me.SuspendLayout()
			' 
			' openFileDialog1
			' 
			Me.openFileDialog1.DefaultExt = "xls"
			Me.openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|Excel 97/2003|*.xls|Excel 2007|*.xlsx;*.xlsm|All " & "files|*.*"
			Me.openFileDialog1.Title = "Open an Excel File"
			' 
			' panel1
			' 
			Me.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
			Me.panel1.Controls.Add(Me.ObjTree)
			Me.panel1.Controls.Add(Me.panel6)
			Me.panel1.Controls.Add(Me.panel3)
			Me.panel1.Dock = System.Windows.Forms.DockStyle.Left
			Me.panel1.Location = New System.Drawing.Point(0, 38)
			Me.panel1.Name = "panel1"
			Me.panel1.Size = New System.Drawing.Size(176, 472)
			Me.panel1.TabIndex = 4
			' 
			' ObjTree
			' 
			Me.ObjTree.Dock = System.Windows.Forms.DockStyle.Fill
			Me.ObjTree.HideSelection = False
			Me.ObjTree.Location = New System.Drawing.Point(0, 112)
			Me.ObjTree.Name = "ObjTree"
			Me.ObjTree.Size = New System.Drawing.Size(172, 356)
			Me.ObjTree.TabIndex = 1
'			Me.ObjTree.AfterSelect += New System.Windows.Forms.TreeViewEventHandler(Me.ObjTree_AfterSelect)
			' 
			' panel6
			' 
			Me.panel6.BackColor = System.Drawing.Color.DarkSeaGreen
			Me.panel6.Controls.Add(Me.cbSheet)
			Me.panel6.Controls.Add(Me.label1)
			Me.panel6.Dock = System.Windows.Forms.DockStyle.Top
			Me.panel6.Location = New System.Drawing.Point(0, 64)
			Me.panel6.Name = "panel6"
			Me.panel6.Size = New System.Drawing.Size(172, 48)
			Me.panel6.TabIndex = 2
			' 
			' cbSheet
			' 
			Me.cbSheet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
			Me.cbSheet.Location = New System.Drawing.Point(6, 16)
			Me.cbSheet.Name = "cbSheet"
			Me.cbSheet.Size = New System.Drawing.Size(160, 21)
			Me.cbSheet.TabIndex = 33
'			Me.cbSheet.SelectedIndexChanged += New System.EventHandler(Me.cbSheet_SelectedIndexChanged)
			' 
			' label1
			' 
			Me.label1.BackColor = System.Drawing.Color.LightGoldenrodYellow
			Me.label1.Dock = System.Windows.Forms.DockStyle.Fill
			Me.label1.ForeColor = System.Drawing.Color.Black
			Me.label1.Location = New System.Drawing.Point(0, 0)
			Me.label1.Name = "label1"
			Me.label1.Size = New System.Drawing.Size(172, 48)
			Me.label1.TabIndex = 1
			Me.label1.Text = "Sheet:"
			' 
			' panel3
			' 
			Me.panel3.BackColor = System.Drawing.Color.DarkSeaGreen
			Me.panel3.Controls.Add(Me.lblObjects)
			Me.panel3.Dock = System.Windows.Forms.DockStyle.Top
			Me.panel3.Location = New System.Drawing.Point(0, 0)
			Me.panel3.Name = "panel3"
			Me.panel3.Size = New System.Drawing.Size(172, 64)
			Me.panel3.TabIndex = 0
			' 
			' lblObjects
			' 
			Me.lblObjects.BackColor = System.Drawing.Color.LightGoldenrodYellow
			Me.lblObjects.Dock = System.Windows.Forms.DockStyle.Fill
			Me.lblObjects.ForeColor = System.Drawing.Color.Black
			Me.lblObjects.Location = New System.Drawing.Point(0, 0)
			Me.lblObjects.Name = "lblObjects"
			Me.lblObjects.Size = New System.Drawing.Size(172, 64)
			Me.lblObjects.TabIndex = 1
			Me.lblObjects.Text = "File:"
			' 
			' splitter1
			' 
			Me.splitter1.Location = New System.Drawing.Point(176, 91)
			Me.splitter1.Name = "splitter1"
			Me.splitter1.Size = New System.Drawing.Size(3, 419)
			Me.splitter1.TabIndex = 5
			Me.splitter1.TabStop = False
			' 
			' dataGrid
			' 
			Me.dataGrid.CaptionText = "Object Properties"
			Me.dataGrid.DataMember = ""
			Me.dataGrid.Dock = System.Windows.Forms.DockStyle.Top
			Me.dataGrid.HeaderForeColor = System.Drawing.SystemColors.ControlText
			Me.dataGrid.Location = New System.Drawing.Point(179, 91)
			Me.dataGrid.Name = "dataGrid"
			Me.dataGrid.PreferredColumnWidth = 120
			Me.dataGrid.ReadOnly = True
			Me.dataGrid.Size = New System.Drawing.Size(557, 213)
			Me.dataGrid.TabIndex = 7
			' 
			' panel4
			' 
			Me.panel4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
			Me.panel4.Controls.Add(Me.PreviewBox)
			Me.panel4.Dock = System.Windows.Forms.DockStyle.Fill
			Me.panel4.Location = New System.Drawing.Point(179, 307)
			Me.panel4.Name = "panel4"
			Me.panel4.Size = New System.Drawing.Size(557, 203)
			Me.panel4.TabIndex = 8
			' 
			' PreviewBox
			' 
			Me.PreviewBox.Dock = System.Windows.Forms.DockStyle.Fill
			Me.PreviewBox.Location = New System.Drawing.Point(0, 0)
			Me.PreviewBox.Name = "PreviewBox"
			Me.PreviewBox.Size = New System.Drawing.Size(553, 199)
			Me.PreviewBox.TabIndex = 0
			Me.PreviewBox.TabStop = False
			' 
			' splitter2
			' 
			Me.splitter2.Dock = System.Windows.Forms.DockStyle.Top
			Me.splitter2.Location = New System.Drawing.Point(179, 304)
			Me.splitter2.Name = "splitter2"
			Me.splitter2.Size = New System.Drawing.Size(557, 3)
			Me.splitter2.TabIndex = 9
			Me.splitter2.TabStop = False
			' 
			' saveImageDialog
			' 
			Me.saveImageDialog.DefaultExt = "png"
			Me.saveImageDialog.Filter = "PNG Files|*.png"
			' 
			' panel5
			' 
			Me.panel5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.panel5.Controls.Add(Me.lblObjText)
			Me.panel5.Controls.Add(Me.lblObjName)
			Me.panel5.Dock = System.Windows.Forms.DockStyle.Top
			Me.panel5.Location = New System.Drawing.Point(176, 38)
			Me.panel5.Name = "panel5"
			Me.panel5.Size = New System.Drawing.Size(560, 53)
			Me.panel5.TabIndex = 10
			' 
			' lblObjText
			' 
			Me.lblObjText.AutoSize = True
			Me.lblObjText.Location = New System.Drawing.Point(16, 32)
			Me.lblObjText.Name = "lblObjText"
			Me.lblObjText.Size = New System.Drawing.Size(31, 13)
			Me.lblObjText.TabIndex = 1
			Me.lblObjText.Text = "Text:"
			' 
			' lblObjName
			' 
			Me.lblObjName.AutoSize = True
			Me.lblObjName.Location = New System.Drawing.Point(16, 8)
			Me.lblObjName.Name = "lblObjName"
			Me.lblObjName.Size = New System.Drawing.Size(38, 13)
			Me.lblObjName.TabIndex = 0
			Me.lblObjName.Text = "Name:"
			' 
			' mainToolbar
			' 
			Me.mainToolbar.Items.AddRange(New System.Windows.Forms.ToolStripItem() { Me.btnOpenFile, Me.toolStripSeparator1, Me.btnShowInExcel, Me.btnSaveAsImage, Me.btnExit, Me.btnInfo, Me.toolStripSeparator2, Me.btnStretchPreview})
			Me.mainToolbar.Location = New System.Drawing.Point(0, 0)
			Me.mainToolbar.Name = "mainToolbar"
			Me.mainToolbar.Size = New System.Drawing.Size(736, 38)
			Me.mainToolbar.TabIndex = 11
			Me.mainToolbar.Text = "toolStrip1"
			' 
			' btnOpenFile
			' 
			Me.btnOpenFile.Image = (CType(resources.GetObject("btnOpenFile.Image"), System.Drawing.Image))
			Me.btnOpenFile.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnOpenFile.Name = "btnOpenFile"
			Me.btnOpenFile.Size = New System.Drawing.Size(59, 35)
			Me.btnOpenFile.Text = "Open file"
			Me.btnOpenFile.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
'			Me.btnOpenFile.Click += New System.EventHandler(Me.btnOpenFile_Click)
			' 
			' toolStripSeparator1
			' 
			Me.toolStripSeparator1.Name = "toolStripSeparator1"
			Me.toolStripSeparator1.Size = New System.Drawing.Size(6, 38)
			' 
			' btnShowInExcel
			' 
			Me.btnShowInExcel.Image = (CType(resources.GetObject("btnShowInExcel.Image"), System.Drawing.Image))
			Me.btnShowInExcel.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnShowInExcel.Name = "btnShowInExcel"
			Me.btnShowInExcel.Size = New System.Drawing.Size(82, 35)
			Me.btnShowInExcel.Text = "Show in Excel"
			Me.btnShowInExcel.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
'			Me.btnShowInExcel.Click += New System.EventHandler(Me.btnOpen_Click)
			' 
			' btnSaveAsImage
			' 
			Me.btnSaveAsImage.Image = (CType(resources.GetObject("btnSaveAsImage.Image"), System.Drawing.Image))
			Me.btnSaveAsImage.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnSaveAsImage.Name = "btnSaveAsImage"
			Me.btnSaveAsImage.Size = New System.Drawing.Size(85, 35)
			Me.btnSaveAsImage.Text = "Save as image"
			Me.btnSaveAsImage.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
'			Me.btnSaveAsImage.Click += New System.EventHandler(Me.btnSaveImage_Click)
			' 
			' btnExit
			' 
			Me.btnExit.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
			Me.btnExit.Image = (CType(resources.GetObject("btnExit.Image"), System.Drawing.Image))
			Me.btnExit.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnExit.Name = "btnExit"
			Me.btnExit.Size = New System.Drawing.Size(59, 35)
			Me.btnExit.Text = "     E&xit     "
			Me.btnExit.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
'			Me.btnExit.Click += New System.EventHandler(Me.btnExit_Click)
			' 
			' btnInfo
			' 
			Me.btnInfo.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
			Me.btnInfo.Image = (CType(resources.GetObject("btnInfo.Image"), System.Drawing.Image))
			Me.btnInfo.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnInfo.Name = "btnInfo"
			Me.btnInfo.Size = New System.Drawing.Size(74, 35)
			Me.btnInfo.Text = "Information"
			Me.btnInfo.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
'			Me.btnInfo.Click += New System.EventHandler(Me.btnInfo_Click)
			' 
			' toolStripSeparator2
			' 
			Me.toolStripSeparator2.Name = "toolStripSeparator2"
			Me.toolStripSeparator2.Size = New System.Drawing.Size(6, 38)
			' 
			' btnStretchPreview
			' 
			Me.btnStretchPreview.CheckOnClick = True
			Me.btnStretchPreview.Image = (CType(resources.GetObject("btnStretchPreview.Image"), System.Drawing.Image))
			Me.btnStretchPreview.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnStretchPreview.Name = "btnStretchPreview"
			Me.btnStretchPreview.Size = New System.Drawing.Size(92, 35)
			Me.btnStretchPreview.Text = "Stretch preview"
			Me.btnStretchPreview.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
'			Me.btnStretchPreview.Click += New System.EventHandler(Me.btnStretchPreview_Click)
			' 
			' mainForm
			' 
			Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New System.Drawing.Size(736, 510)
			Me.Controls.Add(Me.panel4)
			Me.Controls.Add(Me.splitter2)
			Me.Controls.Add(Me.dataGrid)
			Me.Controls.Add(Me.splitter1)
			Me.Controls.Add(Me.panel5)
			Me.Controls.Add(Me.panel1)
			Me.Controls.Add(Me.mainToolbar)
			Me.Name = "mainForm"
			Me.Text = "FlexCel Object Explorer"
			Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
			Me.panel1.ResumeLayout(False)
			Me.panel6.ResumeLayout(False)
			Me.panel3.ResumeLayout(False)
			CType(Me.dataGrid, System.ComponentModel.ISupportInitialize).EndInit()
			Me.panel4.ResumeLayout(False)
			CType(Me.PreviewBox, System.ComponentModel.ISupportInitialize).EndInit()
			Me.panel5.ResumeLayout(False)
			Me.panel5.PerformLayout()
			Me.mainToolbar.ResumeLayout(False)
			Me.mainToolbar.PerformLayout()
			Me.ResumeLayout(False)
			Me.PerformLayout()

		End Sub
		#End Region

		Private mainToolbar As ToolStrip
		Private WithEvents btnOpenFile As ToolStripButton
		Private toolStripSeparator1 As ToolStripSeparator
		Private WithEvents btnShowInExcel As ToolStripButton
		Private WithEvents btnExit As ToolStripButton
		Private WithEvents btnSaveAsImage As ToolStripButton
		Private WithEvents btnInfo As ToolStripButton
		Private toolStripSeparator2 As ToolStripSeparator
		Private WithEvents btnStretchPreview As ToolStripButton


	End Class
End Namespace

