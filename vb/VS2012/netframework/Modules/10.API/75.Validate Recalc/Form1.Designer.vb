Imports System.Collections
Imports System.ComponentModel
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports System.IO
Imports System.Reflection
Namespace ValidateRecalc
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Private openFileDialog1 As System.Windows.Forms.OpenFileDialog
		Private report As System.Windows.Forms.RichTextBox
		Private linkedFileDialog As System.Windows.Forms.OpenFileDialog
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
			Me.report = New System.Windows.Forms.RichTextBox()
			Me.XlsReport = New FlexCel.Report.FlexCelReport()
			Me.linkedFileDialog = New System.Windows.Forms.OpenFileDialog()
			Me.mainToolbar = New System.Windows.Forms.ToolStrip()
			Me.validateRecalc = New System.Windows.Forms.ToolStripButton()
			Me.compareWithExcel = New System.Windows.Forms.ToolStripButton()
			Me.toolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
			Me.btnInfo = New System.Windows.Forms.ToolStripButton()
			Me.btnExit = New System.Windows.Forms.ToolStripButton()
			Me.mainToolbar.SuspendLayout()
			Me.SuspendLayout()
			' 
			' openFileDialog1
			' 
			Me.openFileDialog1.DefaultExt = "xls"
			Me.openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|Excel 97/2003|*.xls|Excel 2007|*.xlsx;*.xlsm|All " & "files|*.*"
			Me.openFileDialog1.Title = "Open an Excel File"
			' 
			' report
			' 
			Me.report.Dock = System.Windows.Forms.DockStyle.Fill
			Me.report.Font = New System.Drawing.Font("Courier New", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.report.Location = New System.Drawing.Point(0, 38)
			Me.report.Name = "report"
			Me.report.ReadOnly = True
			Me.report.Size = New System.Drawing.Size(768, 327)
			Me.report.TabIndex = 3
			Me.report.Text = ""
			' 
			' linkedFileDialog
			' 
			Me.linkedFileDialog.DefaultExt = "xls"
			Me.linkedFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|Excel 97/2003|*.xls|Excel 2007|*.xlsx;*.xlsm|All " & "files|*.*"
			Me.linkedFileDialog.Title = "Please supply the location for the following linked file."
			' 
			' mainToolbar
			' 
			Me.mainToolbar.Items.AddRange(New System.Windows.Forms.ToolStripItem() { Me.validateRecalc, Me.compareWithExcel, Me.toolStripSeparator1, Me.btnInfo, Me.btnExit})
			Me.mainToolbar.Location = New System.Drawing.Point(0, 0)
			Me.mainToolbar.Name = "mainToolbar"
			Me.mainToolbar.Size = New System.Drawing.Size(768, 38)
			Me.mainToolbar.TabIndex = 11
			Me.mainToolbar.Text = "toolStrip1"
			' 
			' validateRecalc
			' 
			Me.validateRecalc.Image = (CType(resources.GetObject("validateRecalc.Image"), System.Drawing.Image))
			Me.validateRecalc.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.validateRecalc.Name = "validateRecalc"
			Me.validateRecalc.Size = New System.Drawing.Size(90, 35)
			Me.validateRecalc.Text = "&Validate Recalc"
			Me.validateRecalc.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
'			Me.validateRecalc.Click += New System.EventHandler(Me.validateRecalc_Click)
			' 
			' compareWithExcel
			' 
			Me.compareWithExcel.Image = (CType(resources.GetObject("compareWithExcel.Image"), System.Drawing.Image))
			Me.compareWithExcel.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.compareWithExcel.Name = "compareWithExcel"
			Me.compareWithExcel.Size = New System.Drawing.Size(115, 43)
			Me.compareWithExcel.Text = "Compare with Excel"
			Me.compareWithExcel.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
'			Me.compareWithExcel.Click += New System.EventHandler(Me.compareWithExcel_Click)
			' 
			' toolStripSeparator1
			' 
			Me.toolStripSeparator1.Name = "toolStripSeparator1"
			Me.toolStripSeparator1.Size = New System.Drawing.Size(6, 46)
			' 
			' btnInfo
			' 
			Me.btnInfo.Image = (CType(resources.GetObject("btnInfo.Image"), System.Drawing.Image))
			Me.btnInfo.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnInfo.Name = "btnInfo"
			Me.btnInfo.Size = New System.Drawing.Size(74, 43)
			Me.btnInfo.Text = "Information"
			Me.btnInfo.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
'			Me.btnInfo.Click += New System.EventHandler(Me.btnInfo_Click)
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
			Me.Controls.Add(Me.report)
			Me.Controls.Add(Me.mainToolbar)
			Me.Name = "mainForm"
			Me.Text = "Validate FlexCel recalculation"
			Me.mainToolbar.ResumeLayout(False)
			Me.mainToolbar.PerformLayout()
			Me.ResumeLayout(False)
			Me.PerformLayout()

		End Sub
		#End Region

		Private mainToolbar As ToolStrip
		Private WithEvents validateRecalc As ToolStripButton
		Private WithEvents compareWithExcel As ToolStripButton
		Private toolStripSeparator1 As ToolStripSeparator
		Private WithEvents btnInfo As ToolStripButton
		Private WithEvents btnExit As ToolStripButton
	End Class
End Namespace

