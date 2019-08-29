Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Namespace FlexCelImageExplorer
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Private openFileDialog1 As System.Windows.Forms.OpenFileDialog
		Private panel1 As System.Windows.Forms.Panel
		Private splitter1 As System.Windows.Forms.Splitter
		Private panel3 As System.Windows.Forms.Panel
		Private WithEvents FilesListBox As System.Windows.Forms.ListBox
		Private lblFolder As System.Windows.Forms.Label
		Private dataGrid As System.Windows.Forms.DataGrid
		Private dataSet1 As System.Data.DataSet
		Private ImageDataTable As System.Data.DataTable
		Private Index As System.Data.DataColumn
		Private Cell1 As System.Data.DataColumn
		Private Cell2 As System.Data.DataColumn
		Private [cType] As System.Data.DataColumn
		Private cText As System.Data.DataColumn
		Private Description As System.Data.DataColumn
		Private dataColumn1 As System.Data.DataColumn
		Private panel4 As System.Windows.Forms.Panel
		Private splitter2 As System.Windows.Forms.Splitter
		Private dataColumn2 As System.Data.DataColumn
		Private PreviewBox As System.Windows.Forms.PictureBox
		Private dataColumn3 As System.Data.DataColumn
		Private saveImageDialog As System.Windows.Forms.SaveFileDialog
		Private dataColumn4 As System.Data.DataColumn
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
			Me.FilesListBox = New System.Windows.Forms.ListBox()
			Me.panel3 = New System.Windows.Forms.Panel()
			Me.lblFolder = New System.Windows.Forms.Label()
			Me.splitter1 = New System.Windows.Forms.Splitter()
			Me.dataGrid = New System.Windows.Forms.DataGrid()
			Me.ImageDataTable = New System.Data.DataTable()
			Me.dataColumn4 = New System.Data.DataColumn()
			Me.Index = New System.Data.DataColumn()
			Me.Cell1 = New System.Data.DataColumn()
			Me.Cell2 = New System.Data.DataColumn()
			Me.cType = New System.Data.DataColumn()
			Me.cText = New System.Data.DataColumn()
			Me.Description = New System.Data.DataColumn()
			Me.dataColumn1 = New System.Data.DataColumn()
			Me.dataColumn2 = New System.Data.DataColumn()
			Me.dataColumn3 = New System.Data.DataColumn()
			Me.dataSet1 = New System.Data.DataSet()
			Me.panel4 = New System.Windows.Forms.Panel()
			Me.PreviewBox = New System.Windows.Forms.PictureBox()
			Me.splitter2 = New System.Windows.Forms.Splitter()
			Me.saveImageDialog = New System.Windows.Forms.SaveFileDialog()
			Me.mainToolbar = New System.Windows.Forms.ToolStrip()
			Me.cbScanFolder = New System.Windows.Forms.ToolStripButton()
			Me.btnOpenFile = New System.Windows.Forms.ToolStripButton()
			Me.toolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
			Me.btnShowInExcel = New System.Windows.Forms.ToolStripButton()
			Me.btnSaveAsImage = New System.Windows.Forms.ToolStripButton()
			Me.btnExit = New System.Windows.Forms.ToolStripButton()
			Me.btnInfo = New System.Windows.Forms.ToolStripButton()
			Me.toolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
			Me.btnStretchPreview = New System.Windows.Forms.ToolStripButton()
			Me.panel1.SuspendLayout()
			Me.panel3.SuspendLayout()
			CType(Me.dataGrid, System.ComponentModel.ISupportInitialize).BeginInit()
			CType(Me.ImageDataTable, System.ComponentModel.ISupportInitialize).BeginInit()
			CType(Me.dataSet1, System.ComponentModel.ISupportInitialize).BeginInit()
			Me.panel4.SuspendLayout()
			CType(Me.PreviewBox, System.ComponentModel.ISupportInitialize).BeginInit()
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
			Me.panel1.Controls.Add(Me.FilesListBox)
			Me.panel1.Controls.Add(Me.panel3)
			Me.panel1.Dock = System.Windows.Forms.DockStyle.Left
			Me.panel1.Location = New System.Drawing.Point(0, 38)
			Me.panel1.Name = "panel1"
			Me.panel1.Size = New System.Drawing.Size(176, 400)
			Me.panel1.TabIndex = 4
			' 
			' FilesListBox
			' 
			Me.FilesListBox.Dock = System.Windows.Forms.DockStyle.Fill
			Me.FilesListBox.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
			Me.FilesListBox.Location = New System.Drawing.Point(0, 40)
			Me.FilesListBox.Name = "FilesListBox"
			Me.FilesListBox.Size = New System.Drawing.Size(172, 356)
			Me.FilesListBox.TabIndex = 1
'			Me.FilesListBox.DrawItem += New System.Windows.Forms.DrawItemEventHandler(Me.FilesListBox_DrawItem)
'			Me.FilesListBox.SelectedIndexChanged += New System.EventHandler(Me.FilesListBox_SelectedIndexChanged)
			' 
			' panel3
			' 
			Me.panel3.BackColor = System.Drawing.Color.DarkSeaGreen
			Me.panel3.Controls.Add(Me.lblFolder)
			Me.panel3.Dock = System.Windows.Forms.DockStyle.Top
			Me.panel3.Location = New System.Drawing.Point(0, 0)
			Me.panel3.Name = "panel3"
			Me.panel3.Size = New System.Drawing.Size(172, 40)
			Me.panel3.TabIndex = 0
			' 
			' lblFolder
			' 
			Me.lblFolder.Dock = System.Windows.Forms.DockStyle.Fill
			Me.lblFolder.ForeColor = System.Drawing.Color.Black
			Me.lblFolder.Location = New System.Drawing.Point(0, 0)
			Me.lblFolder.Name = "lblFolder"
			Me.lblFolder.Size = New System.Drawing.Size(172, 40)
			Me.lblFolder.TabIndex = 0
			Me.lblFolder.Text = "No Selected Folder."
			' 
			' splitter1
			' 
			Me.splitter1.Location = New System.Drawing.Point(176, 38)
			Me.splitter1.Name = "splitter1"
			Me.splitter1.Size = New System.Drawing.Size(3, 400)
			Me.splitter1.TabIndex = 5
			Me.splitter1.TabStop = False
			' 
			' dataGrid
			' 
			Me.dataGrid.CaptionText = "No file selected"
			Me.dataGrid.DataMember = ""
			Me.dataGrid.DataSource = Me.ImageDataTable
			Me.dataGrid.Dock = System.Windows.Forms.DockStyle.Top
			Me.dataGrid.HeaderForeColor = System.Drawing.SystemColors.ControlText
			Me.dataGrid.Location = New System.Drawing.Point(179, 38)
			Me.dataGrid.Name = "dataGrid"
			Me.dataGrid.PreferredColumnWidth = 120
			Me.dataGrid.ReadOnly = True
			Me.dataGrid.Size = New System.Drawing.Size(709, 128)
			Me.dataGrid.TabIndex = 7
			' 
			' ImageDataTable
			' 
			Me.ImageDataTable.Columns.AddRange(New System.Data.DataColumn() { Me.dataColumn4, Me.Index, Me.Cell1, Me.Cell2, Me.cType, Me.cText, Me.Description, Me.dataColumn1, Me.dataColumn2, Me.dataColumn3})
			Me.ImageDataTable.TableName = "ImageDataTable"
			' 
			' dataColumn4
			' 
			Me.dataColumn4.ColumnName = "Sheet"
			' 
			' Index
			' 
			Me.Index.ColumnName = "Index"
			' 
			' Cell1
			' 
			Me.Cell1.Caption = "Width (Pixels)"
			Me.Cell1.ColumnName = "Width (Pixels)"
			' 
			' Cell2
			' 
			Me.Cell2.Caption = "Height (Pixels)"
			Me.Cell2.ColumnName = "Height (Pixels)"
			' 
			' cType
			' 
			Me.cType.ColumnName = "Type"
			' 
			' cText
			' 
			Me.cText.Caption = "Image Format"
			Me.cText.ColumnName = "Image Format"
			' 
			' Description
			' 
			Me.Description.Caption = "Uncompressed size"
			Me.Description.ColumnName = "Uncompressed size"
			' 
			' dataColumn1
			' 
			Me.dataColumn1.Caption = "Name"
			Me.dataColumn1.ColumnName = "Name"
			' 
			' dataColumn2
			' 
			Me.dataColumn2.ColumnMapping = System.Data.MappingType.Hidden
			Me.dataColumn2.ColumnName = "Image"
			Me.dataColumn2.DataType = GetType(Byte())
			' 
			' dataColumn3
			' 
			Me.dataColumn3.ColumnName = "Cropped"
			Me.dataColumn3.DataType = GetType(Boolean)
			' 
			' dataSet1
			' 
			Me.dataSet1.DataSetName = "ImageDataSet"
			Me.dataSet1.Locale = New System.Globalization.CultureInfo("")
			Me.dataSet1.Tables.AddRange(New System.Data.DataTable() { Me.ImageDataTable})
			' 
			' panel4
			' 
			Me.panel4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
			Me.panel4.Controls.Add(Me.PreviewBox)
			Me.panel4.Dock = System.Windows.Forms.DockStyle.Fill
			Me.panel4.Location = New System.Drawing.Point(179, 169)
			Me.panel4.Name = "panel4"
			Me.panel4.Size = New System.Drawing.Size(709, 269)
			Me.panel4.TabIndex = 8
			' 
			' PreviewBox
			' 
			Me.PreviewBox.Dock = System.Windows.Forms.DockStyle.Fill
			Me.PreviewBox.Location = New System.Drawing.Point(0, 0)
			Me.PreviewBox.Name = "PreviewBox"
			Me.PreviewBox.Size = New System.Drawing.Size(705, 265)
			Me.PreviewBox.TabIndex = 0
			Me.PreviewBox.TabStop = False
			' 
			' splitter2
			' 
			Me.splitter2.Dock = System.Windows.Forms.DockStyle.Top
			Me.splitter2.Location = New System.Drawing.Point(179, 166)
			Me.splitter2.Name = "splitter2"
			Me.splitter2.Size = New System.Drawing.Size(709, 3)
			Me.splitter2.TabIndex = 9
			Me.splitter2.TabStop = False
			' 
			' mainToolbar
			' 
			Me.mainToolbar.Items.AddRange(New System.Windows.Forms.ToolStripItem() { Me.cbScanFolder, Me.btnOpenFile, Me.toolStripSeparator1, Me.btnShowInExcel, Me.btnSaveAsImage, Me.btnExit, Me.btnInfo, Me.toolStripSeparator2, Me.btnStretchPreview})
			Me.mainToolbar.Location = New System.Drawing.Point(0, 0)
			Me.mainToolbar.Name = "mainToolbar"
			Me.mainToolbar.Size = New System.Drawing.Size(888, 38)
			Me.mainToolbar.TabIndex = 12
			Me.mainToolbar.Text = "toolStrip1"
			' 
			' cbScanFolder
			' 
			Me.cbScanFolder.CheckOnClick = True
			Me.cbScanFolder.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
			Me.cbScanFolder.Image = (CType(resources.GetObject("cbScanFolder.Image"), System.Drawing.Image))
			Me.cbScanFolder.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.cbScanFolder.Name = "cbScanFolder"
			Me.cbScanFolder.Size = New System.Drawing.Size(122, 35)
			Me.cbScanFolder.Text = "Scan all files in folder"
			Me.cbScanFolder.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
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
			Me.ClientSize = New System.Drawing.Size(888, 438)
			Me.Controls.Add(Me.panel4)
			Me.Controls.Add(Me.splitter2)
			Me.Controls.Add(Me.dataGrid)
			Me.Controls.Add(Me.splitter1)
			Me.Controls.Add(Me.panel1)
			Me.Controls.Add(Me.mainToolbar)
			Me.Name = "mainForm"
			Me.Text = "FlexCel Image Explorer"
			Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
			Me.panel1.ResumeLayout(False)
			Me.panel3.ResumeLayout(False)
			CType(Me.dataGrid, System.ComponentModel.ISupportInitialize).EndInit()
			CType(Me.ImageDataTable, System.ComponentModel.ISupportInitialize).EndInit()
			CType(Me.dataSet1, System.ComponentModel.ISupportInitialize).EndInit()
			Me.panel4.ResumeLayout(False)
			CType(Me.PreviewBox, System.ComponentModel.ISupportInitialize).EndInit()
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
		Private WithEvents btnSaveAsImage As ToolStripButton
		Private WithEvents btnExit As ToolStripButton
		Private WithEvents btnInfo As ToolStripButton
		Private toolStripSeparator2 As ToolStripSeparator
		Private WithEvents btnStretchPreview As ToolStripButton
		Private cbScanFolder As ToolStripButton
	End Class
End Namespace

