Imports System.ComponentModel
Imports System.IO
Imports System.Drawing.Drawing2D
Imports System.Threading
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Winforms
Imports FlexCel.Render
Imports FlexCel.Pdf
Namespace CustomPreview
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		''' <summary>
		''' Required designer variable.
		''' </summary>
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
			Dim resources As New System.ComponentModel.ComponentResourceManager(GetType(mainForm))
			Me.flexCelImgExport1 = New FlexCel.Render.FlexCelImgExport()
			Me.openFileDialog = New System.Windows.Forms.OpenFileDialog()
			Me.panel1 = New System.Windows.Forms.Panel()
			Me.MainPreview = New FlexCel.Winforms.FlexCelPreview()
			Me.thumbs = New FlexCel.Winforms.FlexCelPreview()
			Me.splitter1 = New System.Windows.Forms.Splitter()
			Me.panelLeft = New System.Windows.Forms.Panel()
			Me.cbAllSheets = New System.Windows.Forms.CheckBox()
			Me.label2 = New System.Windows.Forms.Label()
			Me.sheetSplitter = New System.Windows.Forms.Splitter()
			Me.lbSheets = New System.Windows.Forms.ListBox()
			Me.label1 = New System.Windows.Forms.Label()
			Me.PdfSaveFileDialog = New System.Windows.Forms.SaveFileDialog()
			Me.toolTip1 = New System.Windows.Forms.ToolTip(Me.components)
			Me.mainToolbar = New System.Windows.Forms.ToolStrip()
			Me.openFile = New System.Windows.Forms.ToolStripButton()
			Me.toolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
			Me.btnFirst = New System.Windows.Forms.ToolStripButton()
			Me.btnPrev = New System.Windows.Forms.ToolStripButton()
			Me.edPage = New System.Windows.Forms.ToolStripTextBox()
			Me.btnNext = New System.Windows.Forms.ToolStripButton()
			Me.btnLast = New System.Windows.Forms.ToolStripButton()
			Me.toolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
			Me.btnAutofit = New System.Windows.Forms.ToolStripDropDownButton()
			Me.noneToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
			Me.fitToWidthToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
			Me.fitToHeightToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
			Me.fitToPageToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
			Me.btnZoomOut = New System.Windows.Forms.ToolStripButton()
			Me.edZoom = New System.Windows.Forms.ToolStripTextBox()
			Me.btnZoomIn = New System.Windows.Forms.ToolStripButton()
			Me.toolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
			Me.btnGridLines = New System.Windows.Forms.ToolStripButton()
			Me.btnHeadings = New System.Windows.Forms.ToolStripButton()
			Me.toolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
			Me.btnRecalc = New System.Windows.Forms.ToolStripButton()
			Me.btnPdf = New System.Windows.Forms.ToolStripButton()
			Me.btnClose = New System.Windows.Forms.ToolStripButton()
			Me.panel1.SuspendLayout()
			Me.panelLeft.SuspendLayout()
			Me.mainToolbar.SuspendLayout()
			Me.SuspendLayout()
			' 
			' flexCelImgExport1
			' 
			Me.flexCelImgExport1.AllVisibleSheets = False
			Me.flexCelImgExport1.PageSize = Nothing
			Me.flexCelImgExport1.ResetPageNumberOnEachSheet = False
			Me.flexCelImgExport1.Resolution = 96R
			Me.flexCelImgExport1.Workbook = Nothing
			' 
			' openFileDialog
			' 
			Me.openFileDialog.DefaultExt = "xls"
			Me.openFileDialog.Filter = "Excel Files|*.xls|All files|*.*"
			Me.openFileDialog.Title = "Select a file to preview"
			' 
			' panel1
			' 
			Me.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.panel1.Controls.Add(Me.MainPreview)
			Me.panel1.Controls.Add(Me.splitter1)
			Me.panel1.Controls.Add(Me.panelLeft)
			Me.panel1.Dock = System.Windows.Forms.DockStyle.Fill
			Me.panel1.Location = New System.Drawing.Point(0, 46)
			Me.panel1.Name = "panel1"
			Me.panel1.Size = New System.Drawing.Size(808, 375)
			Me.panel1.TabIndex = 8
			' 
			' MainPreview
			' 
			Me.MainPreview.AutoScrollMinSize = New System.Drawing.Size(40, 383)
			Me.MainPreview.Dock = System.Windows.Forms.DockStyle.Fill
			Me.MainPreview.Document = Me.flexCelImgExport1
			Me.MainPreview.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic
			Me.MainPreview.Location = New System.Drawing.Point(144, 0)
			Me.MainPreview.Name = "MainPreview"
			Me.MainPreview.PageXSeparation = 20
			Me.MainPreview.Size = New System.Drawing.Size(662, 373)
			Me.MainPreview.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias
			Me.MainPreview.StartPage = 1
			Me.MainPreview.TabIndex = 2
			Me.MainPreview.ThumbnailLarge = Nothing
			Me.MainPreview.ThumbnailSmall = Me.thumbs
'			Me.MainPreview.StartPageChanged += New System.EventHandler(Me.flexCelPreview1_StartPageChanged)
'			Me.MainPreview.ZoomChanged += New System.EventHandler(Me.flexCelPreview1_ZoomChanged)
			' 
			' thumbs
			' 
			Me.thumbs.AutoScrollMinSize = New System.Drawing.Size(20, 10)
			Me.thumbs.Dock = System.Windows.Forms.DockStyle.Fill
			Me.thumbs.Document = Me.flexCelImgExport1
			Me.thumbs.Location = New System.Drawing.Point(0, 115)
			Me.thumbs.Name = "thumbs"
			Me.thumbs.Size = New System.Drawing.Size(136, 258)
			Me.thumbs.StartPage = 1
			Me.thumbs.TabIndex = 3
			Me.thumbs.ThumbnailLarge = Me.MainPreview
			Me.thumbs.ThumbnailSmall = Nothing
			Me.thumbs.Zoom = 0.1R
			' 
			' splitter1
			' 
			Me.splitter1.BackColor = System.Drawing.SystemColors.ControlLightLight
			Me.splitter1.Location = New System.Drawing.Point(136, 0)
			Me.splitter1.MinSize = 0
			Me.splitter1.Name = "splitter1"
			Me.splitter1.Size = New System.Drawing.Size(8, 373)
			Me.splitter1.TabIndex = 11
			Me.splitter1.TabStop = False
			' 
			' panelLeft
			' 
			Me.panelLeft.Controls.Add(Me.cbAllSheets)
			Me.panelLeft.Controls.Add(Me.thumbs)
			Me.panelLeft.Controls.Add(Me.label2)
			Me.panelLeft.Controls.Add(Me.sheetSplitter)
			Me.panelLeft.Controls.Add(Me.lbSheets)
			Me.panelLeft.Controls.Add(Me.label1)
			Me.panelLeft.Dock = System.Windows.Forms.DockStyle.Left
			Me.panelLeft.Location = New System.Drawing.Point(0, 0)
			Me.panelLeft.Name = "panelLeft"
			Me.panelLeft.Size = New System.Drawing.Size(136, 373)
			Me.panelLeft.TabIndex = 9
			' 
			' cbAllSheets
			' 
			Me.cbAllSheets.Location = New System.Drawing.Point(16, 16)
			Me.cbAllSheets.Name = "cbAllSheets"
			Me.cbAllSheets.Size = New System.Drawing.Size(104, 16)
			Me.cbAllSheets.TabIndex = 14
			Me.cbAllSheets.Text = "All Sheets"
'			Me.cbAllSheets.CheckedChanged += New System.EventHandler(Me.cbAllSheets_CheckedChanged)
			' 
			' label2
			' 
			Me.label2.Dock = System.Windows.Forms.DockStyle.Top
			Me.label2.Location = New System.Drawing.Point(0, 99)
			Me.label2.Name = "label2"
			Me.label2.Size = New System.Drawing.Size(136, 16)
			Me.label2.TabIndex = 13
			Me.label2.Text = "Thumbs"
			' 
			' sheetSplitter
			' 
			Me.sheetSplitter.BackColor = System.Drawing.SystemColors.ControlLightLight
			Me.sheetSplitter.Dock = System.Windows.Forms.DockStyle.Top
			Me.sheetSplitter.Location = New System.Drawing.Point(0, 91)
			Me.sheetSplitter.Name = "sheetSplitter"
			Me.sheetSplitter.Size = New System.Drawing.Size(136, 8)
			Me.sheetSplitter.TabIndex = 11
			Me.sheetSplitter.TabStop = False
			' 
			' lbSheets
			' 
			Me.lbSheets.Dock = System.Windows.Forms.DockStyle.Top
			Me.lbSheets.Items.AddRange(New Object() { "No open file"})
			Me.lbSheets.Location = New System.Drawing.Point(0, 35)
			Me.lbSheets.Name = "lbSheets"
			Me.lbSheets.Size = New System.Drawing.Size(136, 56)
			Me.lbSheets.TabIndex = 10
'			Me.lbSheets.SelectedIndexChanged += New System.EventHandler(Me.lbSheets_SelectedIndexChanged)
			' 
			' label1
			' 
			Me.label1.Dock = System.Windows.Forms.DockStyle.Top
			Me.label1.Location = New System.Drawing.Point(0, 0)
			Me.label1.Name = "label1"
			Me.label1.Size = New System.Drawing.Size(136, 35)
			Me.label1.TabIndex = 12
			Me.label1.Text = "Sheets"
			' 
			' PdfSaveFileDialog
			' 
			Me.PdfSaveFileDialog.DefaultExt = "pdf"
			Me.PdfSaveFileDialog.Filter = "Pdf Files|*.pdf"
			Me.PdfSaveFileDialog.Title = "Select the file to export to:"
			' 
			' mainToolbar
			' 
			Me.mainToolbar.ImageScalingSize = New System.Drawing.Size(24, 24)
			Me.mainToolbar.Items.AddRange(New System.Windows.Forms.ToolStripItem() { Me.openFile, Me.toolStripSeparator2, Me.btnFirst, Me.btnPrev, Me.edPage, Me.btnNext, Me.btnLast, Me.toolStripSeparator1, Me.btnAutofit, Me.btnZoomOut, Me.edZoom, Me.btnZoomIn, Me.toolStripSeparator3, Me.btnGridLines, Me.btnHeadings, Me.toolStripSeparator4, Me.btnRecalc, Me.btnPdf, Me.btnClose})
			Me.mainToolbar.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.HorizontalStackWithOverflow
			Me.mainToolbar.Location = New System.Drawing.Point(0, 0)
			Me.mainToolbar.Name = "mainToolbar"
			Me.mainToolbar.Size = New System.Drawing.Size(808, 46)
			Me.mainToolbar.TabIndex = 14
			' 
			' openFile
			' 
			Me.openFile.Image = My.Resources.open
			Me.openFile.ImageAlign = System.Drawing.ContentAlignment.TopCenter
			Me.openFile.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.openFile.Name = "openFile"
			Me.openFile.Size = New System.Drawing.Size(61, 43)
			Me.openFile.Text = "&Open File"
			Me.openFile.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
			Me.openFile.ToolTipText = "Open an Excel file"
'			Me.openFile.Click += New System.EventHandler(Me.openFile_Click)
			' 
			' toolStripSeparator2
			' 
			Me.toolStripSeparator2.AutoSize = False
			Me.toolStripSeparator2.Name = "toolStripSeparator2"
			Me.toolStripSeparator2.Size = New System.Drawing.Size(20, 46)
			' 
			' btnFirst
			' 
			Me.btnFirst.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
			Me.btnFirst.Enabled = False
			Me.btnFirst.Image = (CType(resources.GetObject("btnFirst.Image"), System.Drawing.Image))
			Me.btnFirst.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnFirst.Name = "btnFirst"
			Me.btnFirst.Size = New System.Drawing.Size(27, 43)
			Me.btnFirst.Text = "<<"
			Me.btnFirst.ToolTipText = "First page"
'			Me.btnFirst.Click += New System.EventHandler(Me.btnFirst_Click)
			' 
			' btnPrev
			' 
			Me.btnPrev.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
			Me.btnPrev.Enabled = False
			Me.btnPrev.Image = (CType(resources.GetObject("btnPrev.Image"), System.Drawing.Image))
			Me.btnPrev.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnPrev.Name = "btnPrev"
			Me.btnPrev.Size = New System.Drawing.Size(23, 43)
			Me.btnPrev.Text = "<"
			Me.btnPrev.ToolTipText = "Previous page"
'			Me.btnPrev.Click += New System.EventHandler(Me.btnPrev_Click)
			' 
			' edPage
			' 
			Me.edPage.AutoSize = False
			Me.edPage.Enabled = False
			Me.edPage.Name = "edPage"
			Me.edPage.Size = New System.Drawing.Size(100, 18)
			Me.edPage.TextBoxTextAlign = System.Windows.Forms.HorizontalAlignment.Right
'			Me.edPage.Leave += New System.EventHandler(Me.edPage_Leave)
'			Me.edPage.KeyPress += New System.Windows.Forms.KeyPressEventHandler(Me.edPage_KeyPress)
			' 
			' btnNext
			' 
			Me.btnNext.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
			Me.btnNext.Enabled = False
			Me.btnNext.Image = (CType(resources.GetObject("btnNext.Image"), System.Drawing.Image))
			Me.btnNext.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnNext.Name = "btnNext"
			Me.btnNext.Size = New System.Drawing.Size(23, 43)
			Me.btnNext.Text = ">"
			Me.btnNext.ToolTipText = "Next page"
'			Me.btnNext.Click += New System.EventHandler(Me.btnNext_Click)
			' 
			' btnLast
			' 
			Me.btnLast.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
			Me.btnLast.Enabled = False
			Me.btnLast.Image = (CType(resources.GetObject("btnLast.Image"), System.Drawing.Image))
			Me.btnLast.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnLast.Name = "btnLast"
			Me.btnLast.Size = New System.Drawing.Size(27, 43)
			Me.btnLast.Text = ">>"
			Me.btnLast.ToolTipText = "Last page"
'			Me.btnLast.Click += New System.EventHandler(Me.btnLast_Click)
			' 
			' toolStripSeparator1
			' 
			Me.toolStripSeparator1.Name = "toolStripSeparator1"
			Me.toolStripSeparator1.Size = New System.Drawing.Size(6, 46)
			' 
			' btnAutofit
			' 
			Me.btnAutofit.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() { Me.noneToolStripMenuItem, Me.fitToWidthToolStripMenuItem, Me.fitToHeightToolStripMenuItem, Me.fitToPageToolStripMenuItem})
			Me.btnAutofit.Image = My.Resources.autofit
			Me.btnAutofit.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnAutofit.Name = "btnAutofit"
			Me.btnAutofit.Size = New System.Drawing.Size(76, 43)
			Me.btnAutofit.Text = "No Autofit"
			Me.btnAutofit.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
			' 
			' noneToolStripMenuItem
			' 
			Me.noneToolStripMenuItem.Name = "noneToolStripMenuItem"
			Me.noneToolStripMenuItem.Size = New System.Drawing.Size(140, 22)
			Me.noneToolStripMenuItem.Text = "No Autofit"
'			Me.noneToolStripMenuItem.Click += New System.EventHandler(Me.noneToolStripMenuItem_Click)
			' 
			' fitToWidthToolStripMenuItem
			' 
			Me.fitToWidthToolStripMenuItem.Name = "fitToWidthToolStripMenuItem"
			Me.fitToWidthToolStripMenuItem.Size = New System.Drawing.Size(140, 22)
			Me.fitToWidthToolStripMenuItem.Text = "Fit to Width"
'			Me.fitToWidthToolStripMenuItem.Click += New System.EventHandler(Me.fitToWidthToolStripMenuItem_Click)
			' 
			' fitToHeightToolStripMenuItem
			' 
			Me.fitToHeightToolStripMenuItem.Name = "fitToHeightToolStripMenuItem"
			Me.fitToHeightToolStripMenuItem.Size = New System.Drawing.Size(140, 22)
			Me.fitToHeightToolStripMenuItem.Text = "Fit to Height"
'			Me.fitToHeightToolStripMenuItem.Click += New System.EventHandler(Me.fitToHeightToolStripMenuItem_Click)
			' 
			' fitToPageToolStripMenuItem
			' 
			Me.fitToPageToolStripMenuItem.Name = "fitToPageToolStripMenuItem"
			Me.fitToPageToolStripMenuItem.Size = New System.Drawing.Size(140, 22)
			Me.fitToPageToolStripMenuItem.Text = "Fit to Page"
'			Me.fitToPageToolStripMenuItem.Click += New System.EventHandler(Me.fitToPageToolStripMenuItem_Click)
			' 
			' btnZoomOut
			' 
			Me.btnZoomOut.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
			Me.btnZoomOut.Enabled = False
			Me.btnZoomOut.Image = (CType(resources.GetObject("btnZoomOut.Image"), System.Drawing.Image))
			Me.btnZoomOut.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnZoomOut.Name = "btnZoomOut"
			Me.btnZoomOut.Size = New System.Drawing.Size(23, 43)
			Me.btnZoomOut.Text = "-"
			Me.btnZoomOut.ToolTipText = "Zoom out"
'			Me.btnZoomOut.Click += New System.EventHandler(Me.btnZoomOut_Click)
			' 
			' edZoom
			' 
			Me.edZoom.AutoSize = False
			Me.edZoom.Enabled = False
			Me.edZoom.Name = "edZoom"
			Me.edZoom.Size = New System.Drawing.Size(40, 18)
			Me.edZoom.TextBoxTextAlign = System.Windows.Forms.HorizontalAlignment.Right
'			Me.edZoom.Enter += New System.EventHandler(Me.edZoom_Enter)
'			Me.edZoom.KeyPress += New System.Windows.Forms.KeyPressEventHandler(Me.edZoom_KeyPress)
			' 
			' btnZoomIn
			' 
			Me.btnZoomIn.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
			Me.btnZoomIn.Enabled = False
			Me.btnZoomIn.Image = (CType(resources.GetObject("btnZoomIn.Image"), System.Drawing.Image))
			Me.btnZoomIn.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnZoomIn.Name = "btnZoomIn"
			Me.btnZoomIn.Size = New System.Drawing.Size(23, 43)
			Me.btnZoomIn.Text = "+"
			Me.btnZoomIn.ToolTipText = "Zoom in"
'			Me.btnZoomIn.Click += New System.EventHandler(Me.btnZoomIn_Click)
			' 
			' toolStripSeparator3
			' 
			Me.toolStripSeparator3.AutoSize = False
			Me.toolStripSeparator3.Name = "toolStripSeparator3"
			Me.toolStripSeparator3.Size = New System.Drawing.Size(20, 46)
			' 
			' btnGridLines
			' 
			Me.btnGridLines.CheckOnClick = True
			Me.btnGridLines.Enabled = False
			Me.btnGridLines.Image = My.Resources.grid
			Me.btnGridLines.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnGridLines.Name = "btnGridLines"
			Me.btnGridLines.Size = New System.Drawing.Size(57, 43)
			Me.btnGridLines.Text = "&Gridlines"
			Me.btnGridLines.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
			Me.btnGridLines.ToolTipText = "Show gridlines"
'			Me.btnGridLines.Click += New System.EventHandler(Me.btnGridLines_Click)
			' 
			' btnHeadings
			' 
			Me.btnHeadings.CheckOnClick = True
			Me.btnHeadings.Enabled = False
			Me.btnHeadings.Image = My.Resources.Head
			Me.btnHeadings.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnHeadings.Name = "btnHeadings"
			Me.btnHeadings.Size = New System.Drawing.Size(61, 43)
			Me.btnHeadings.Text = "&Headings"
			Me.btnHeadings.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
			Me.btnHeadings.ToolTipText = "Show the headings"
'			Me.btnHeadings.Click += New System.EventHandler(Me.btnHeadings_Click)
			' 
			' toolStripSeparator4
			' 
			Me.toolStripSeparator4.Name = "toolStripSeparator4"
			Me.toolStripSeparator4.Size = New System.Drawing.Size(6, 46)
			' 
			' btnRecalc
			' 
			Me.btnRecalc.Enabled = False
			Me.btnRecalc.Image = My.Resources.calc
			Me.btnRecalc.ImageAlign = System.Drawing.ContentAlignment.TopCenter
			Me.btnRecalc.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnRecalc.Name = "btnRecalc"
			Me.btnRecalc.Size = New System.Drawing.Size(45, 43)
			Me.btnRecalc.Text = "&Recalc"
			Me.btnRecalc.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
			Me.btnRecalc.ToolTipText = "Recalculate the file"
'			Me.btnRecalc.Click += New System.EventHandler(Me.btnRecalc_Click)
			' 
			' btnPdf
			' 
			Me.btnPdf.Enabled = False
			Me.btnPdf.Image = My.Resources.pdf
			Me.btnPdf.ImageAlign = System.Drawing.ContentAlignment.TopCenter
			Me.btnPdf.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnPdf.Name = "btnPdf"
			Me.btnPdf.Size = New System.Drawing.Size(79, 43)
			Me.btnPdf.Text = "Export to &Pdf"
			Me.btnPdf.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
			Me.btnPdf.ToolTipText = "Export the file to Pdf"
'			Me.btnPdf.Click += New System.EventHandler(Me.btnPdf_Click)
			' 
			' btnClose
			' 
			Me.btnClose.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
			Me.btnClose.Image = My.Resources.close
			Me.btnClose.ImageAlign = System.Drawing.ContentAlignment.TopCenter
			Me.btnClose.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnClose.Name = "btnClose"
			Me.btnClose.Size = New System.Drawing.Size(59, 43)
			Me.btnClose.Text = "     E&xit     "
			Me.btnClose.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
			Me.btnClose.ToolTipText = "Exit from the application"
'			Me.btnClose.Click += New System.EventHandler(Me.btnClose_Click)
			' 
			' mainForm
			' 
			Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New System.Drawing.Size(808, 421)
			Me.Controls.Add(Me.panel1)
			Me.Controls.Add(Me.mainToolbar)
			Me.Name = "mainForm"
			Me.Text = "Custom Preview Demo"
'			Me.Load += New System.EventHandler(Me.mainForm_Load)
			Me.panel1.ResumeLayout(False)
			Me.panelLeft.ResumeLayout(False)
			Me.mainToolbar.ResumeLayout(False)
			Me.mainToolbar.PerformLayout()
			Me.ResumeLayout(False)
			Me.PerformLayout()

		End Sub
		#End Region

		Private openFileDialog As System.Windows.Forms.OpenFileDialog
		Private panel1 As System.Windows.Forms.Panel
		Private panelLeft As System.Windows.Forms.Panel
		Private WithEvents lbSheets As System.Windows.Forms.ListBox
		Private splitter1 As System.Windows.Forms.Splitter
		Private label1 As System.Windows.Forms.Label
		Private label2 As System.Windows.Forms.Label
		Private PdfSaveFileDialog As System.Windows.Forms.SaveFileDialog
		Private WithEvents cbAllSheets As System.Windows.Forms.CheckBox
		Private sheetSplitter As System.Windows.Forms.Splitter
		Private toolTip1 As System.Windows.Forms.ToolTip
		Private flexCelImgExport1 As FlexCel.Render.FlexCelImgExport
		Private WithEvents MainPreview As FlexCel.Winforms.FlexCelPreview
		Private thumbs As FlexCel.Winforms.FlexCelPreview
		Private mainToolbar As ToolStrip
		Private WithEvents openFile As ToolStripButton
		Private toolStripSeparator1 As ToolStripSeparator
		Private WithEvents btnRecalc As ToolStripButton
		Private WithEvents btnPdf As ToolStripButton
		Private WithEvents btnClose As ToolStripButton
		Private toolStripSeparator2 As ToolStripSeparator
		Private WithEvents btnFirst As ToolStripButton
		Private WithEvents btnPrev As ToolStripButton
		Private WithEvents edPage As ToolStripTextBox
		Private WithEvents btnNext As ToolStripButton
		Private WithEvents btnLast As ToolStripButton
		Private WithEvents btnZoomOut As ToolStripButton
		Private WithEvents edZoom As ToolStripTextBox
		Private WithEvents btnZoomIn As ToolStripButton
		Private toolStripSeparator3 As ToolStripSeparator
		Private WithEvents btnHeadings As ToolStripButton
		Private WithEvents btnGridLines As ToolStripButton
		Private toolStripSeparator4 As ToolStripSeparator
		Private btnAutofit As ToolStripDropDownButton
		Private WithEvents noneToolStripMenuItem As ToolStripMenuItem
		Private WithEvents fitToWidthToolStripMenuItem As ToolStripMenuItem
		Private WithEvents fitToHeightToolStripMenuItem As ToolStripMenuItem
		Private WithEvents fitToPageToolStripMenuItem As ToolStripMenuItem
	End Class
End Namespace

