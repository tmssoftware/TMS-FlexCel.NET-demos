Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports System.Data.OleDb
Imports System.Threading
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Report

Namespace GenericReports2
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Private saveFileDialog1 As System.Windows.Forms.SaveFileDialog
		Private dataSet As System.Data.DataSet
		Private dataGrid As System.Windows.Forms.DataGrid
		Private Connection As System.Data.OleDb.OleDbConnection
		Private dbDataAdapter As System.Data.OleDb.OleDbDataAdapter
		Private Report As FlexCel.Report.FlexCelReport
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
			Me.Report = New FlexCel.Report.FlexCelReport()
			Me.Connection = New System.Data.OleDb.OleDbConnection()
			Me.dataSet = New System.Data.DataSet()
			Me.dataGrid = New System.Windows.Forms.DataGrid()
			Me.dbDataAdapter = New System.Data.OleDb.OleDbDataAdapter()
			Me.mainToolbar = New System.Windows.Forms.ToolStrip()
			Me.btnOpenConnection = New System.Windows.Forms.ToolStripButton()
			Me.btnQuery = New System.Windows.Forms.ToolStripButton()
			Me.toolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
			Me.btnExportExcel = New System.Windows.Forms.ToolStripButton()
			Me.btnExit = New System.Windows.Forms.ToolStripButton()
			CType(Me.dataSet, System.ComponentModel.ISupportInitialize).BeginInit()
			CType(Me.dataGrid, System.ComponentModel.ISupportInitialize).BeginInit()
			Me.mainToolbar.SuspendLayout()
			Me.SuspendLayout()
			' 
			' saveFileDialog1
			' 
			Me.saveFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|Excel 97/2003|*.xls|Excel 2007|*.xlsx;*.xlsm|All " & "files|*.*"
			Me.saveFileDialog1.RestoreDirectory = True
			' 
			' Report
			' 
			Me.Report.AllowOverwritingFiles = True
			Me.Report.DeleteEmptyBands = FlexCel.Report.TDeleteEmptyBands.ClearDataAndFormats
			Me.Report.DeleteEmptyRanges = False
			' 
			' dataSet
			' 
			Me.dataSet.DataSetName = "NewDataSet"
			Me.dataSet.Locale = New System.Globalization.CultureInfo("es-ES")
			' 
			' dataGrid
			' 
			Me.dataGrid.DataMember = ""
			Me.dataGrid.Dock = System.Windows.Forms.DockStyle.Fill
			Me.dataGrid.HeaderForeColor = System.Drawing.SystemColors.ControlText
			Me.dataGrid.Location = New System.Drawing.Point(0, 38)
			Me.dataGrid.Name = "dataGrid"
			Me.dataGrid.Size = New System.Drawing.Size(528, 239)
			Me.dataGrid.TabIndex = 4
			' 
			' mainToolbar
			' 
			Me.mainToolbar.Items.AddRange(New System.Windows.Forms.ToolStripItem() { Me.btnOpenConnection, Me.btnQuery, Me.toolStripSeparator1, Me.btnExportExcel, Me.btnExit})
			Me.mainToolbar.Location = New System.Drawing.Point(0, 0)
			Me.mainToolbar.Name = "mainToolbar"
			Me.mainToolbar.Size = New System.Drawing.Size(528, 38)
			Me.mainToolbar.TabIndex = 11
			Me.mainToolbar.Text = "mainToolbar"
			' 
			' btnOpenConnection
			' 
			Me.btnOpenConnection.Image = (CType(resources.GetObject("btnOpenConnection.Image"), System.Drawing.Image))
			Me.btnOpenConnection.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnOpenConnection.Name = "btnOpenConnection"
			Me.btnOpenConnection.Size = New System.Drawing.Size(103, 35)
			Me.btnOpenConnection.Text = "Open connection"
			Me.btnOpenConnection.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
'			Me.btnOpenConnection.Click += New System.EventHandler(Me.btnOpenconnection_Click)
			' 
			' btnQuery
			' 
			Me.btnQuery.Image = (CType(resources.GetObject("btnQuery.Image"), System.Drawing.Image))
			Me.btnQuery.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnQuery.Name = "btnQuery"
			Me.btnQuery.Size = New System.Drawing.Size(70, 35)
			Me.btnQuery.Text = "Query Data"
			Me.btnQuery.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
'			Me.btnQuery.Click += New System.EventHandler(Me.btnQuery_Click)
			' 
			' toolStripSeparator1
			' 
			Me.toolStripSeparator1.Name = "toolStripSeparator1"
			Me.toolStripSeparator1.Size = New System.Drawing.Size(6, 38)
			' 
			' btnExportExcel
			' 
			Me.btnExportExcel.Image = (CType(resources.GetObject("btnExportExcel.Image"), System.Drawing.Image))
			Me.btnExportExcel.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnExportExcel.Name = "btnExportExcel"
			Me.btnExportExcel.Size = New System.Drawing.Size(87, 35)
			Me.btnExportExcel.Text = "Export to Excel"
			Me.btnExportExcel.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
'			Me.btnExportExcel.Click += New System.EventHandler(Me.btnExportExcel_Click)
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
			Me.ClientSize = New System.Drawing.Size(528, 277)
			Me.Controls.Add(Me.dataGrid)
			Me.Controls.Add(Me.mainToolbar)
			Me.Name = "mainForm"
			Me.Text = "Generic Reports 2"
			CType(Me.dataSet, System.ComponentModel.ISupportInitialize).EndInit()
			CType(Me.dataGrid, System.ComponentModel.ISupportInitialize).EndInit()
			Me.mainToolbar.ResumeLayout(False)
			Me.mainToolbar.PerformLayout()
			Me.ResumeLayout(False)
			Me.PerformLayout()

		End Sub
		#End Region

		Private mainToolbar As ToolStrip
		Private WithEvents btnOpenConnection As ToolStripButton
		Private WithEvents btnQuery As ToolStripButton
		Private toolStripSeparator1 As ToolStripSeparator
		Private WithEvents btnExportExcel As ToolStripButton
		Private WithEvents btnExit As ToolStripButton
	End Class
End Namespace

