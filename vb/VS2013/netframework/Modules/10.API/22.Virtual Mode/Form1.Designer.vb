Imports System.Collections
Imports System.ComponentModel
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Namespace VirtualMode
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Private openFileDialog1 As System.Windows.Forms.OpenFileDialog
		Private panel2 As System.Windows.Forms.Panel
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
			Dim dataGridViewCellStyle1 As New System.Windows.Forms.DataGridViewCellStyle()
			Dim resources As New System.ComponentModel.ComponentResourceManager(GetType(mainForm))
			Me.openFileDialog1 = New System.Windows.Forms.OpenFileDialog()
			Me.panel2 = New System.Windows.Forms.Panel()
			Me.cbFormatValues = New System.Windows.Forms.CheckBox()
			Me.cbIgnoreFormulaText = New System.Windows.Forms.CheckBox()
			Me.cbFirst50Rows = New System.Windows.Forms.CheckBox()
			Me.statusBar = New System.Windows.Forms.StatusBar()
			Me.DisplayGrid = New System.Windows.Forms.DataGridView()
			Me.GridCaptionPanel = New System.Windows.Forms.Panel()
			Me.GridCaption = New System.Windows.Forms.Label()
			Me.mainToolbar = New System.Windows.Forms.ToolStrip()
			Me.toolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
			Me.btnOpenFile = New System.Windows.Forms.ToolStripButton()
			Me.btnExit = New System.Windows.Forms.ToolStripButton()
			Me.btnInfo = New System.Windows.Forms.ToolStripButton()
			Me.panel2.SuspendLayout()
			CType(Me.DisplayGrid, System.ComponentModel.ISupportInitialize).BeginInit()
			Me.GridCaptionPanel.SuspendLayout()
			Me.mainToolbar.SuspendLayout()
			Me.SuspendLayout()
			' 
			' openFileDialog1
			' 
			Me.openFileDialog1.DefaultExt = "xls"
			Me.openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|Excel 97/2003|*.xls|Excel 2007|*.xlsx;*.xlsm|All " & "files|*.*"
			Me.openFileDialog1.Title = "Open an Excel File"
			' 
			' panel2
			' 
			Me.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.panel2.Controls.Add(Me.cbFormatValues)
			Me.panel2.Controls.Add(Me.cbIgnoreFormulaText)
			Me.panel2.Controls.Add(Me.cbFirst50Rows)
			Me.panel2.Dock = System.Windows.Forms.DockStyle.Top
			Me.panel2.Location = New System.Drawing.Point(0, 38)
			Me.panel2.Name = "panel2"
			Me.panel2.Size = New System.Drawing.Size(880, 25)
			Me.panel2.TabIndex = 4
			' 
			' cbFormatValues
			' 
			Me.cbFormatValues.AutoSize = True
			Me.cbFormatValues.Location = New System.Drawing.Point(269, 5)
			Me.cbFormatValues.Name = "cbFormatValues"
			Me.cbFormatValues.Size = New System.Drawing.Size(131, 17)
			Me.cbFormatValues.TabIndex = 2
			Me.cbFormatValues.Text = "Format values (slower)"
			Me.cbFormatValues.UseVisualStyleBackColor = True
			' 
			' cbIgnoreFormulaText
			' 
			Me.cbIgnoreFormulaText.AutoSize = True
			Me.cbIgnoreFormulaText.Checked = True
			Me.cbIgnoreFormulaText.CheckState = System.Windows.Forms.CheckState.Checked
			Me.cbIgnoreFormulaText.Location = New System.Drawing.Point(150, 5)
			Me.cbIgnoreFormulaText.Name = "cbIgnoreFormulaText"
			Me.cbIgnoreFormulaText.Size = New System.Drawing.Size(113, 17)
			Me.cbIgnoreFormulaText.TabIndex = 1
			Me.cbIgnoreFormulaText.Text = "Ignore formula text"
			Me.cbIgnoreFormulaText.UseVisualStyleBackColor = True
			' 
			' cbFirst50Rows
			' 
			Me.cbFirst50Rows.AutoSize = True
			Me.cbFirst50Rows.Checked = True
			Me.cbFirst50Rows.CheckState = System.Windows.Forms.CheckState.Checked
			Me.cbFirst50Rows.Location = New System.Drawing.Point(11, 5)
			Me.cbFirst50Rows.Name = "cbFirst50Rows"
			Me.cbFirst50Rows.Size = New System.Drawing.Size(133, 17)
			Me.cbFirst50Rows.TabIndex = 0
			Me.cbFirst50Rows.Text = "Read only first 50 rows"
			Me.cbFirst50Rows.UseVisualStyleBackColor = True
			' 
			' statusBar
			' 
			Me.statusBar.Location = New System.Drawing.Point(0, 439)
			Me.statusBar.Name = "statusBar"
			Me.statusBar.Size = New System.Drawing.Size(880, 22)
			Me.statusBar.TabIndex = 7
			' 
			' DisplayGrid
			' 
			Me.DisplayGrid.AllowUserToAddRows = False
			Me.DisplayGrid.AllowUserToDeleteRows = False
			dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
			dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
			dataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
			dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
			dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
			dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True
			Me.DisplayGrid.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1
			Me.DisplayGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
			Me.DisplayGrid.Dock = System.Windows.Forms.DockStyle.Fill
			Me.DisplayGrid.Location = New System.Drawing.Point(0, 86)
			Me.DisplayGrid.Name = "DisplayGrid"
			Me.DisplayGrid.ReadOnly = True
			Me.DisplayGrid.Size = New System.Drawing.Size(880, 353)
			Me.DisplayGrid.TabIndex = 8
			Me.DisplayGrid.VirtualMode = True
'			Me.DisplayGrid.CellValueNeeded += New System.Windows.Forms.DataGridViewCellValueEventHandler(Me.DisplayGrid_CellValueNeeded)
'			Me.DisplayGrid.RowPostPaint += New System.Windows.Forms.DataGridViewRowPostPaintEventHandler(Me.DisplayGrid_RowPostPaint)
			' 
			' GridCaptionPanel
			' 
			Me.GridCaptionPanel.BackColor = System.Drawing.SystemColors.ActiveCaption
			Me.GridCaptionPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.GridCaptionPanel.Controls.Add(Me.GridCaption)
			Me.GridCaptionPanel.Dock = System.Windows.Forms.DockStyle.Top
			Me.GridCaptionPanel.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
			Me.GridCaptionPanel.Location = New System.Drawing.Point(0, 63)
			Me.GridCaptionPanel.Name = "GridCaptionPanel"
			Me.GridCaptionPanel.Size = New System.Drawing.Size(880, 23)
			Me.GridCaptionPanel.TabIndex = 9
			' 
			' GridCaption
			' 
			Me.GridCaption.AutoSize = True
			Me.GridCaption.Location = New System.Drawing.Point(13, 6)
			Me.GridCaption.Name = "GridCaption"
			Me.GridCaption.Size = New System.Drawing.Size(0, 13)
			Me.GridCaption.TabIndex = 0
			' 
			' mainToolbar
			' 
			Me.mainToolbar.Items.AddRange(New System.Windows.Forms.ToolStripItem() { Me.btnOpenFile, Me.toolStripSeparator1, Me.btnExit, Me.btnInfo})
			Me.mainToolbar.Location = New System.Drawing.Point(0, 0)
			Me.mainToolbar.Name = "mainToolbar"
			Me.mainToolbar.Size = New System.Drawing.Size(880, 38)
			Me.mainToolbar.TabIndex = 10
			Me.mainToolbar.Text = "toolStrip1"
			' 
			' toolStripSeparator1
			' 
			Me.toolStripSeparator1.Name = "toolStripSeparator1"
			Me.toolStripSeparator1.Size = New System.Drawing.Size(6, 38)
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
			Me.Controls.Add(Me.GridCaptionPanel)
			Me.Controls.Add(Me.panel2)
			Me.Controls.Add(Me.statusBar)
			Me.Controls.Add(Me.mainToolbar)
			Me.Name = "mainForm"
			Me.Text = "Virtual Mode Example - Cells are not stored in memory"
			Me.panel2.ResumeLayout(False)
			Me.panel2.PerformLayout()
			CType(Me.DisplayGrid, System.ComponentModel.ISupportInitialize).EndInit()
			Me.GridCaptionPanel.ResumeLayout(False)
			Me.GridCaptionPanel.PerformLayout()
			Me.mainToolbar.ResumeLayout(False)
			Me.mainToolbar.PerformLayout()
			Me.ResumeLayout(False)
			Me.PerformLayout()

		End Sub
		#End Region

		Private WithEvents DisplayGrid As DataGridView
		Private GridCaptionPanel As Panel
		Private GridCaption As Label
		Private mainToolbar As ToolStrip
		Private WithEvents btnOpenFile As ToolStripButton
		Private toolStripSeparator1 As ToolStripSeparator
		Private WithEvents btnInfo As ToolStripButton
		Private WithEvents btnExit As ToolStripButton
		Private cbFirst50Rows As CheckBox
		Private cbIgnoreFormulaText As CheckBox
		Private cbFormatValues As CheckBox
	End Class
End Namespace

