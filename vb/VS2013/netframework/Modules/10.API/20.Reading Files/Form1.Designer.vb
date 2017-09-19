Imports System.Collections
Imports System.ComponentModel
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Namespace ReadingFiles
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Private openFileDialog1 As System.Windows.Forms.OpenFileDialog
		Private DisplayGrid As System.Windows.Forms.DataGrid
		Private panel1 As System.Windows.Forms.Panel
		Private label1 As System.Windows.Forms.Label
		Private WithEvents sheetCombo As System.Windows.Forms.ComboBox
		Private statusBar As System.Windows.Forms.StatusBar
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
			Me.DisplayGrid = New System.Windows.Forms.DataGrid()
			Me.panel1 = New System.Windows.Forms.Panel()
			Me.sheetCombo = New System.Windows.Forms.ComboBox()
			Me.label1 = New System.Windows.Forms.Label()
			Me.statusBar = New System.Windows.Forms.StatusBar()
			Me.mainToolbar = New System.Windows.Forms.ToolStrip()
			Me.btnOpenFile = New System.Windows.Forms.ToolStripButton()
			Me.toolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
			Me.btnFormatValues = New System.Windows.Forms.ToolStripButton()
			Me.toolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
			Me.btnValueInCellA1 = New System.Windows.Forms.ToolStripButton()
			Me.btnExit = New System.Windows.Forms.ToolStripButton()
			Me.btnInfo = New System.Windows.Forms.ToolStripButton()
			CType(Me.DisplayGrid, System.ComponentModel.ISupportInitialize).BeginInit()
			Me.panel1.SuspendLayout()
			Me.mainToolbar.SuspendLayout()
			Me.SuspendLayout()
			' 
			' openFileDialog1
			' 
			Me.openFileDialog1.DefaultExt = "xls"
			Me.openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|Excel 97/2003|*.xls|Excel 2007|*.xlsx;*.xlsm|All " & "files|*.*"
			Me.openFileDialog1.Title = "Open an Excel File"
			' 
			' DisplayGrid
			' 
			Me.DisplayGrid.DataMember = ""
			Me.DisplayGrid.Dock = System.Windows.Forms.DockStyle.Fill
			Me.DisplayGrid.HeaderForeColor = System.Drawing.SystemColors.ControlText
			Me.DisplayGrid.Location = New System.Drawing.Point(0, 67)
			Me.DisplayGrid.Name = "DisplayGrid"
			Me.DisplayGrid.Size = New System.Drawing.Size(880, 372)
			Me.DisplayGrid.TabIndex = 5
			' 
			' panel1
			' 
			Me.panel1.BackColor = System.Drawing.SystemColors.ControlDark
			Me.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.panel1.Controls.Add(Me.sheetCombo)
			Me.panel1.Controls.Add(Me.label1)
			Me.panel1.Dock = System.Windows.Forms.DockStyle.Top
			Me.panel1.Location = New System.Drawing.Point(0, 38)
			Me.panel1.Name = "panel1"
			Me.panel1.Size = New System.Drawing.Size(880, 29)
			Me.panel1.TabIndex = 6
			' 
			' sheetCombo
			' 
			Me.sheetCombo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
			Me.sheetCombo.Location = New System.Drawing.Point(65, 3)
			Me.sheetCombo.Name = "sheetCombo"
			Me.sheetCombo.Size = New System.Drawing.Size(391, 21)
			Me.sheetCombo.TabIndex = 1
'			Me.sheetCombo.SelectedIndexChanged += New System.EventHandler(Me.sheetCombo_SelectedIndexChanged)
			' 
			' label1
			' 
			Me.label1.ForeColor = System.Drawing.SystemColors.HighlightText
			Me.label1.Location = New System.Drawing.Point(8, 8)
			Me.label1.Name = "label1"
			Me.label1.Size = New System.Drawing.Size(40, 23)
			Me.label1.TabIndex = 0
			Me.label1.Text = "Sheet:"
			' 
			' statusBar
			' 
			Me.statusBar.Location = New System.Drawing.Point(0, 439)
			Me.statusBar.Name = "statusBar"
			Me.statusBar.Size = New System.Drawing.Size(880, 22)
			Me.statusBar.TabIndex = 7
			' 
			' mainToolbar
			' 
			Me.mainToolbar.Items.AddRange(New System.Windows.Forms.ToolStripItem() { Me.btnOpenFile, Me.toolStripSeparator1, Me.btnFormatValues, Me.toolStripSeparator2, Me.btnValueInCellA1, Me.btnExit, Me.btnInfo})
			Me.mainToolbar.Location = New System.Drawing.Point(0, 0)
			Me.mainToolbar.Name = "mainToolbar"
			Me.mainToolbar.Size = New System.Drawing.Size(880, 38)
			Me.mainToolbar.TabIndex = 11
			Me.mainToolbar.Text = "mainToolbar"
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
			' btnFormatValues
			' 
			Me.btnFormatValues.CheckOnClick = True
			Me.btnFormatValues.Image = (CType(resources.GetObject("btnFormatValues.Image"), System.Drawing.Image))
			Me.btnFormatValues.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnFormatValues.Name = "btnFormatValues"
			Me.btnFormatValues.Size = New System.Drawing.Size(85, 35)
			Me.btnFormatValues.Text = "&Format values"
			Me.btnFormatValues.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
			' 
			' toolStripSeparator2
			' 
			Me.toolStripSeparator2.Name = "toolStripSeparator2"
			Me.toolStripSeparator2.Size = New System.Drawing.Size(6, 38)
			' 
			' btnValueInCellA1
			' 
			Me.btnValueInCellA1.Image = (CType(resources.GetObject("btnValueInCellA1.Image"), System.Drawing.Image))
			Me.btnValueInCellA1.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnValueInCellA1.Name = "btnValueInCellA1"
			Me.btnValueInCellA1.Size = New System.Drawing.Size(91, 35)
			Me.btnValueInCellA1.Text = "&Value in cell A1"
			Me.btnValueInCellA1.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
'			Me.btnValueInCellA1.Click += New System.EventHandler(Me.btnValueInCurrentCell_Click)
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
			' mainForm
			' 
			Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New System.Drawing.Size(880, 461)
			Me.Controls.Add(Me.DisplayGrid)
			Me.Controls.Add(Me.panel1)
			Me.Controls.Add(Me.statusBar)
			Me.Controls.Add(Me.mainToolbar)
			Me.Name = "mainForm"
			Me.Text = "Reading Excel Files"
			CType(Me.DisplayGrid, System.ComponentModel.ISupportInitialize).EndInit()
			Me.panel1.ResumeLayout(False)
			Me.mainToolbar.ResumeLayout(False)
			Me.mainToolbar.PerformLayout()
			Me.ResumeLayout(False)
			Me.PerformLayout()

		End Sub
		#End Region

		Private mainToolbar As ToolStrip
		Private WithEvents btnOpenFile As ToolStripButton
		Private toolStripSeparator1 As ToolStripSeparator
		Private WithEvents btnInfo As ToolStripButton
		Private WithEvents btnExit As ToolStripButton
		Private WithEvents btnValueInCellA1 As ToolStripButton
		Private toolStripSeparator2 As ToolStripSeparator
		Private btnFormatValues As ToolStripButton
	End Class
End Namespace

