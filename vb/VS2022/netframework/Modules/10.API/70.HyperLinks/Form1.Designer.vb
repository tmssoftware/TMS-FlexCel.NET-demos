Imports System.Collections
Imports System.ComponentModel
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports System.IO
Imports System.Reflection
Namespace HyperLinks
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Private saveFileDialog1 As System.Windows.Forms.SaveFileDialog
		Private openFileDialog1 As System.Windows.Forms.OpenFileDialog
		Private dataGrid As System.Windows.Forms.DataGrid
		Private dataSet1 As System.Data.DataSet
		Private HlDataTable As System.Data.DataTable
		Private Index As System.Data.DataColumn
		Private Cell1 As System.Data.DataColumn
		Private Cell2 As System.Data.DataColumn
		Private [cType] As System.Data.DataColumn
		Private Description As System.Data.DataColumn
		Private TextMark As System.Data.DataColumn
		Private TargetFrame As System.Data.DataColumn
		Private cText As System.Data.DataColumn
		Private cHint As System.Data.DataColumn
		Private dataGridTableStyle1 As System.Windows.Forms.DataGridTableStyle
		Private dataGridTextBoxColumn1 As System.Windows.Forms.DataGridTextBoxColumn
		Private dataGridTextBoxColumn2 As System.Windows.Forms.DataGridTextBoxColumn
		Private dataGridTextBoxColumn3 As System.Windows.Forms.DataGridTextBoxColumn
		Private dataGridTextBoxColumn4 As System.Windows.Forms.DataGridTextBoxColumn
		Private dataGridTextBoxColumn5 As System.Windows.Forms.DataGridTextBoxColumn
		Private dataGridTextBoxColumn6 As System.Windows.Forms.DataGridTextBoxColumn
		Private dataGridTextBoxColumn7 As System.Windows.Forms.DataGridTextBoxColumn
		Private dataGridTextBoxColumn8 As System.Windows.Forms.DataGridTextBoxColumn
		Private dataGridTextBoxColumn9 As System.Windows.Forms.DataGridTextBoxColumn
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
			Me.saveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
			Me.dataGrid = New System.Windows.Forms.DataGrid()
			Me.HlDataTable = New System.Data.DataTable()
			Me.Index = New System.Data.DataColumn()
			Me.Cell1 = New System.Data.DataColumn()
			Me.Cell2 = New System.Data.DataColumn()
			Me.cType = New System.Data.DataColumn()
			Me.cText = New System.Data.DataColumn()
			Me.Description = New System.Data.DataColumn()
			Me.TextMark = New System.Data.DataColumn()
			Me.TargetFrame = New System.Data.DataColumn()
			Me.cHint = New System.Data.DataColumn()
			Me.dataGridTableStyle1 = New System.Windows.Forms.DataGridTableStyle()
			Me.dataGridTextBoxColumn1 = New System.Windows.Forms.DataGridTextBoxColumn()
			Me.dataGridTextBoxColumn2 = New System.Windows.Forms.DataGridTextBoxColumn()
			Me.dataGridTextBoxColumn3 = New System.Windows.Forms.DataGridTextBoxColumn()
			Me.dataGridTextBoxColumn4 = New System.Windows.Forms.DataGridTextBoxColumn()
			Me.dataGridTextBoxColumn5 = New System.Windows.Forms.DataGridTextBoxColumn()
			Me.dataGridTextBoxColumn6 = New System.Windows.Forms.DataGridTextBoxColumn()
			Me.dataGridTextBoxColumn7 = New System.Windows.Forms.DataGridTextBoxColumn()
			Me.dataGridTextBoxColumn8 = New System.Windows.Forms.DataGridTextBoxColumn()
			Me.dataGridTextBoxColumn9 = New System.Windows.Forms.DataGridTextBoxColumn()
			Me.openFileDialog1 = New System.Windows.Forms.OpenFileDialog()
			Me.dataSet1 = New System.Data.DataSet()
			Me.mainToolbar = New System.Windows.Forms.ToolStrip()
			Me.btnReadHyperlinks = New System.Windows.Forms.ToolStripButton()
			Me.btnWriteHyperlinks = New System.Windows.Forms.ToolStripButton()
			Me.btnExit = New System.Windows.Forms.ToolStripButton()
			CType(Me.dataGrid, System.ComponentModel.ISupportInitialize).BeginInit()
			CType(Me.HlDataTable, System.ComponentModel.ISupportInitialize).BeginInit()
			CType(Me.dataSet1, System.ComponentModel.ISupportInitialize).BeginInit()
			Me.mainToolbar.SuspendLayout()
			Me.SuspendLayout()
			' 
			' saveFileDialog1
			' 
			Me.saveFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|Excel 97/2003|*.xls|Excel 2007|*.xlsx;*.xlsm|All " & "files|*.*"
			Me.saveFileDialog1.RestoreDirectory = True
			Me.saveFileDialog1.Title = "Save the file as..."
			' 
			' dataGrid
			' 
			Me.dataGrid.CaptionText = "No file selected"
			Me.dataGrid.DataMember = ""
			Me.dataGrid.DataSource = Me.HlDataTable
			Me.dataGrid.Dock = System.Windows.Forms.DockStyle.Fill
			Me.dataGrid.HeaderForeColor = System.Drawing.SystemColors.ControlText
			Me.dataGrid.Location = New System.Drawing.Point(0, 38)
			Me.dataGrid.Name = "dataGrid"
			Me.dataGrid.ReadOnly = True
			Me.dataGrid.Size = New System.Drawing.Size(768, 327)
			Me.dataGrid.TabIndex = 3
			Me.dataGrid.TableStyles.AddRange(New System.Windows.Forms.DataGridTableStyle() { Me.dataGridTableStyle1})
			' 
			' HlDataTable
			' 
			Me.HlDataTable.Columns.AddRange(New System.Data.DataColumn() { Me.Index, Me.Cell1, Me.Cell2, Me.cType, Me.cText, Me.Description, Me.TextMark, Me.TargetFrame, Me.cHint})
			Me.HlDataTable.TableName = "HlDataTable"
			' 
			' Index
			' 
			Me.Index.ColumnName = "Index"
			' 
			' Cell1
			' 
			Me.Cell1.ColumnName = "Cell1"
			' 
			' Cell2
			' 
			Me.Cell2.ColumnName = "Cell2"
			' 
			' cType
			' 
			Me.cType.ColumnName = "Type"
			' 
			' cText
			' 
			Me.cText.ColumnName = "Text"
			' 
			' Description
			' 
			Me.Description.ColumnName = "Description"
			' 
			' TextMark
			' 
			Me.TextMark.ColumnName = "TextMark"
			' 
			' TargetFrame
			' 
			Me.TargetFrame.ColumnName = "TargetFrame"
			' 
			' cHint
			' 
			Me.cHint.ColumnName = "Hint"
			' 
			' dataGridTableStyle1
			' 
			Me.dataGridTableStyle1.DataGrid = Me.dataGrid
			Me.dataGridTableStyle1.GridColumnStyles.AddRange(New System.Windows.Forms.DataGridColumnStyle() { Me.dataGridTextBoxColumn1, Me.dataGridTextBoxColumn2, Me.dataGridTextBoxColumn3, Me.dataGridTextBoxColumn4, Me.dataGridTextBoxColumn5, Me.dataGridTextBoxColumn6, Me.dataGridTextBoxColumn7, Me.dataGridTextBoxColumn8, Me.dataGridTextBoxColumn9})
			Me.dataGridTableStyle1.HeaderForeColor = System.Drawing.SystemColors.ControlText
			Me.dataGridTableStyle1.MappingName = "HlDataTable"
			Me.dataGridTableStyle1.PreferredColumnWidth = 15
			' 
			' dataGridTextBoxColumn1
			' 
			Me.dataGridTextBoxColumn1.Format = ""
			Me.dataGridTextBoxColumn1.FormatInfo = Nothing
			Me.dataGridTextBoxColumn1.HeaderText = "Index"
			Me.dataGridTextBoxColumn1.MappingName = "Index"
			Me.dataGridTextBoxColumn1.Width = 35
			' 
			' dataGridTextBoxColumn2
			' 
			Me.dataGridTextBoxColumn2.Format = ""
			Me.dataGridTextBoxColumn2.FormatInfo = Nothing
			Me.dataGridTextBoxColumn2.HeaderText = "Cell1"
			Me.dataGridTextBoxColumn2.MappingName = "Cell1"
			Me.dataGridTextBoxColumn2.Width = 40
			' 
			' dataGridTextBoxColumn3
			' 
			Me.dataGridTextBoxColumn3.Format = ""
			Me.dataGridTextBoxColumn3.FormatInfo = Nothing
			Me.dataGridTextBoxColumn3.HeaderText = "Cell2"
			Me.dataGridTextBoxColumn3.MappingName = "Cell2"
			Me.dataGridTextBoxColumn3.Width = 40
			' 
			' dataGridTextBoxColumn4
			' 
			Me.dataGridTextBoxColumn4.Format = ""
			Me.dataGridTextBoxColumn4.FormatInfo = Nothing
			Me.dataGridTextBoxColumn4.HeaderText = "Type"
			Me.dataGridTextBoxColumn4.MappingName = "Type"
			Me.dataGridTextBoxColumn4.Width = 75
			' 
			' dataGridTextBoxColumn5
			' 
			Me.dataGridTextBoxColumn5.Format = ""
			Me.dataGridTextBoxColumn5.FormatInfo = Nothing
			Me.dataGridTextBoxColumn5.HeaderText = "Text"
			Me.dataGridTextBoxColumn5.MappingName = "Text"
			Me.dataGridTextBoxColumn5.Width = 150
			' 
			' dataGridTextBoxColumn6
			' 
			Me.dataGridTextBoxColumn6.Format = ""
			Me.dataGridTextBoxColumn6.FormatInfo = Nothing
			Me.dataGridTextBoxColumn6.HeaderText = "Description"
			Me.dataGridTextBoxColumn6.MappingName = "Description"
			Me.dataGridTextBoxColumn6.Width = 150
			' 
			' dataGridTextBoxColumn7
			' 
			Me.dataGridTextBoxColumn7.Format = ""
			Me.dataGridTextBoxColumn7.FormatInfo = Nothing
			Me.dataGridTextBoxColumn7.HeaderText = "TextMark"
			Me.dataGridTextBoxColumn7.MappingName = "TextMark"
			Me.dataGridTextBoxColumn7.Width = 75
			' 
			' dataGridTextBoxColumn8
			' 
			Me.dataGridTextBoxColumn8.Format = ""
			Me.dataGridTextBoxColumn8.FormatInfo = Nothing
			Me.dataGridTextBoxColumn8.HeaderText = "TargetFrame"
			Me.dataGridTextBoxColumn8.MappingName = "TargetFrame"
			Me.dataGridTextBoxColumn8.Width = 75
			' 
			' dataGridTextBoxColumn9
			' 
			Me.dataGridTextBoxColumn9.Format = ""
			Me.dataGridTextBoxColumn9.FormatInfo = Nothing
			Me.dataGridTextBoxColumn9.HeaderText = "Hint"
			Me.dataGridTextBoxColumn9.MappingName = "Hint"
			Me.dataGridTextBoxColumn9.Width = 75
			' 
			' openFileDialog1
			' 
			Me.openFileDialog1.DefaultExt = "xls"
			Me.openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|Excel 97/2003|*.xls|Excel 2007|*.xlsx;*.xlsm|All " & "files|*.*"
			Me.openFileDialog1.Title = "Open an Excel File"
			' 
			' dataSet1
			' 
			Me.dataSet1.DataSetName = "HlDataSet"
			Me.dataSet1.Locale = New System.Globalization.CultureInfo("")
			Me.dataSet1.Tables.AddRange(New System.Data.DataTable() { Me.HlDataTable})
			' 
			' mainToolbar
			' 
			Me.mainToolbar.Items.AddRange(New System.Windows.Forms.ToolStripItem() { Me.btnReadHyperlinks, Me.btnWriteHyperlinks, Me.btnExit})
			Me.mainToolbar.Location = New System.Drawing.Point(0, 0)
			Me.mainToolbar.Name = "mainToolbar"
			Me.mainToolbar.Size = New System.Drawing.Size(768, 38)
			Me.mainToolbar.TabIndex = 11
			Me.mainToolbar.Text = "toolStrip1"
			' 
			' btnReadHyperlinks
			' 
			Me.btnReadHyperlinks.Image = (CType(resources.GetObject("btnReadHyperlinks.Image"), System.Drawing.Image))
			Me.btnReadHyperlinks.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnReadHyperlinks.Name = "btnReadHyperlinks"
			Me.btnReadHyperlinks.Size = New System.Drawing.Size(96, 35)
			Me.btnReadHyperlinks.Text = "Read Hyperlinks"
			Me.btnReadHyperlinks.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
'			Me.btnReadHyperlinks.Click += New System.EventHandler(Me.ReadHyperLinks_Click)
			' 
			' btnWriteHyperlinks
			' 
			Me.btnWriteHyperlinks.Image = (CType(resources.GetObject("btnWriteHyperlinks.Image"), System.Drawing.Image))
			Me.btnWriteHyperlinks.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnWriteHyperlinks.Name = "btnWriteHyperlinks"
			Me.btnWriteHyperlinks.Size = New System.Drawing.Size(98, 43)
			Me.btnWriteHyperlinks.Text = "Write Hyperlinks"
			Me.btnWriteHyperlinks.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
'			Me.btnWriteHyperlinks.Click += New System.EventHandler(Me.writeHyperLinks_Click)
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
'			Me.btnExit.Click += New System.EventHandler(Me.button2_Click)
			' 
			' mainForm
			' 
			Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New System.Drawing.Size(768, 365)
			Me.Controls.Add(Me.dataGrid)
			Me.Controls.Add(Me.mainToolbar)
			Me.Name = "mainForm"
			Me.Text = "Form1"
			CType(Me.dataGrid, System.ComponentModel.ISupportInitialize).EndInit()
			CType(Me.HlDataTable, System.ComponentModel.ISupportInitialize).EndInit()
			CType(Me.dataSet1, System.ComponentModel.ISupportInitialize).EndInit()
			Me.mainToolbar.ResumeLayout(False)
			Me.mainToolbar.PerformLayout()
			Me.ResumeLayout(False)
			Me.PerformLayout()

		End Sub
		#End Region

		Private mainToolbar As ToolStrip
		Private WithEvents btnReadHyperlinks As ToolStripButton
		Private WithEvents btnWriteHyperlinks As ToolStripButton
		Private WithEvents btnExit As ToolStripButton
	End Class
End Namespace

