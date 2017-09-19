Imports System.Drawing.Imaging
Imports System.Collections
Imports System.ComponentModel
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Render
Imports System.IO
Imports System.Reflection
Imports System.Drawing.Drawing2D
Imports System.Drawing.Printing
Imports System.Runtime.InteropServices
Namespace PrintPreviewandExport
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Private openFileDialog1 As System.Windows.Forms.OpenFileDialog
		Private printDialog1 As System.Windows.Forms.PrintDialog
		Private printPreviewDialog1 As System.Windows.Forms.PrintPreviewDialog
		Private panel1 As System.Windows.Forms.Panel
		Private label1 As System.Windows.Forms.Label
		Private edFileName As System.Windows.Forms.TextBox
		Private chFormulaText As System.Windows.Forms.CheckBox
		Private chAntiAlias As System.Windows.Forms.CheckBox
		Private chGridLines As System.Windows.Forms.CheckBox
		Private edHeader As System.Windows.Forms.TextBox
		Private label2 As System.Windows.Forms.Label
		Private edFooter As System.Windows.Forms.TextBox
		Private label3 As System.Windows.Forms.Label
		Private edHPages As System.Windows.Forms.TextBox
		Private edVPages As System.Windows.Forms.TextBox
		Private label5 As System.Windows.Forms.Label
		Private label6 As System.Windows.Forms.Label
		Private chPrintLeft As System.Windows.Forms.CheckBox
		Private WithEvents chFitIn As System.Windows.Forms.CheckBox
		Private label4 As System.Windows.Forms.Label
		Private edZoom As System.Windows.Forms.TextBox
		Private label7 As System.Windows.Forms.Label
		Private edl As System.Windows.Forms.TextBox
		Private label8 As System.Windows.Forms.Label
		Private edt As System.Windows.Forms.TextBox
		Private label9 As System.Windows.Forms.Label
		Private edr As System.Windows.Forms.TextBox
		Private labelb As System.Windows.Forms.Label
		Private edb As System.Windows.Forms.TextBox
		Private label10 As System.Windows.Forms.Label
		Private edh As System.Windows.Forms.TextBox
		Private label11 As System.Windows.Forms.Label
		Private edf As System.Windows.Forms.TextBox
		Private Landscape As System.Windows.Forms.CheckBox
		Private panel3 As System.Windows.Forms.Panel
		Private label12 As System.Windows.Forms.Label
		Private label13 As System.Windows.Forms.Label
		Private label14 As System.Windows.Forms.Label
		Private edTop As System.Windows.Forms.TextBox
		Private edLeft As System.Windows.Forms.TextBox
		Private label15 As System.Windows.Forms.Label
		Private edRight As System.Windows.Forms.TextBox
		Private label16 As System.Windows.Forms.Label
		Private edBottom As System.Windows.Forms.TextBox
		Private label17 As System.Windows.Forms.Label
		Private WithEvents cbSheet As System.Windows.Forms.ComboBox
		Private label18 As System.Windows.Forms.Label
		Private cbConfidential As System.Windows.Forms.CheckBox
		Private exportImageDialog As System.Windows.Forms.SaveFileDialog
		Private chHeadings As System.Windows.Forms.CheckBox
		Private label19 As System.Windows.Forms.Label
		Private cbInterpolation As System.Windows.Forms.ComboBox
		Private exportTiffDialog As System.Windows.Forms.SaveFileDialog
		Private WithEvents cbAllSheets As System.Windows.Forms.CheckBox
		Private panel4 As System.Windows.Forms.Panel
		Private cbResetPageNumber As System.Windows.Forms.CheckBox
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
			Me.printDialog1 = New System.Windows.Forms.PrintDialog()
			Me.flexCelPrintDocument1 = New FlexCel.Render.FlexCelPrintDocument()
			Me.printPreviewDialog1 = New System.Windows.Forms.PrintPreviewDialog()
			Me.panel1 = New System.Windows.Forms.Panel()
			Me.cbResetPageNumber = New System.Windows.Forms.CheckBox()
			Me.panel4 = New System.Windows.Forms.Panel()
			Me.cbAllSheets = New System.Windows.Forms.CheckBox()
			Me.label19 = New System.Windows.Forms.Label()
			Me.cbInterpolation = New System.Windows.Forms.ComboBox()
			Me.chHeadings = New System.Windows.Forms.CheckBox()
			Me.cbConfidential = New System.Windows.Forms.CheckBox()
			Me.label18 = New System.Windows.Forms.Label()
			Me.cbSheet = New System.Windows.Forms.ComboBox()
			Me.panel3 = New System.Windows.Forms.Panel()
			Me.edBottom = New System.Windows.Forms.TextBox()
			Me.label17 = New System.Windows.Forms.Label()
			Me.edRight = New System.Windows.Forms.TextBox()
			Me.label16 = New System.Windows.Forms.Label()
			Me.edLeft = New System.Windows.Forms.TextBox()
			Me.label15 = New System.Windows.Forms.Label()
			Me.edTop = New System.Windows.Forms.TextBox()
			Me.label14 = New System.Windows.Forms.Label()
			Me.label13 = New System.Windows.Forms.Label()
			Me.label12 = New System.Windows.Forms.Label()
			Me.Landscape = New System.Windows.Forms.CheckBox()
			Me.label11 = New System.Windows.Forms.Label()
			Me.edf = New System.Windows.Forms.TextBox()
			Me.labelb = New System.Windows.Forms.Label()
			Me.edb = New System.Windows.Forms.TextBox()
			Me.label9 = New System.Windows.Forms.Label()
			Me.edr = New System.Windows.Forms.TextBox()
			Me.label8 = New System.Windows.Forms.Label()
			Me.edt = New System.Windows.Forms.TextBox()
			Me.label7 = New System.Windows.Forms.Label()
			Me.edl = New System.Windows.Forms.TextBox()
			Me.label4 = New System.Windows.Forms.Label()
			Me.edZoom = New System.Windows.Forms.TextBox()
			Me.chFitIn = New System.Windows.Forms.CheckBox()
			Me.chPrintLeft = New System.Windows.Forms.CheckBox()
			Me.label6 = New System.Windows.Forms.Label()
			Me.label5 = New System.Windows.Forms.Label()
			Me.edVPages = New System.Windows.Forms.TextBox()
			Me.edHPages = New System.Windows.Forms.TextBox()
			Me.edFooter = New System.Windows.Forms.TextBox()
			Me.label3 = New System.Windows.Forms.Label()
			Me.edHeader = New System.Windows.Forms.TextBox()
			Me.label2 = New System.Windows.Forms.Label()
			Me.edFileName = New System.Windows.Forms.TextBox()
			Me.chFormulaText = New System.Windows.Forms.CheckBox()
			Me.chGridLines = New System.Windows.Forms.CheckBox()
			Me.chAntiAlias = New System.Windows.Forms.CheckBox()
			Me.label1 = New System.Windows.Forms.Label()
			Me.label10 = New System.Windows.Forms.Label()
			Me.edh = New System.Windows.Forms.TextBox()
			Me.exportImageDialog = New System.Windows.Forms.SaveFileDialog()
			Me.exportTiffDialog = New System.Windows.Forms.SaveFileDialog()
			Me.mainToolbar = New System.Windows.Forms.ToolStrip()
			Me.btnOpenFile = New System.Windows.Forms.ToolStripButton()
			Me.toolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
			Me.btnSetup = New System.Windows.Forms.ToolStripButton()
			Me.btnPreview = New System.Windows.Forms.ToolStripButton()
			Me.btnPrint = New System.Windows.Forms.ToolStripButton()
			Me.toolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
			Me.btnExportAsImages = New System.Windows.Forms.ToolStripDropDownButton()
			Me.usingFlexCelImgExportAllPagesrecommendedWayToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
			Me.blackAndWhiteToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
			Me.colorsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
			Me.trueColorToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
			Me.usingFlexCelImgExport1PagerecommendedWayToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
			Me.blackAndWhiteToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
			Me.colorsToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
			Me.trueColorToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
			Me.multiPageTIFFUsingFlexCelImgExportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
			Me.faxToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
			Me.blackAndWhiteToolStripMenuItem2 = New System.Windows.Forms.ToolStripMenuItem()
			Me.colorsToolStripMenuItem2 = New System.Windows.Forms.ToolStripMenuItem()
			Me.trueColorToolStripMenuItem2 = New System.Windows.Forms.ToolStripMenuItem()
			Me.btnExit = New System.Windows.Forms.ToolStripButton()
			Me.panel1.SuspendLayout()
			Me.panel3.SuspendLayout()
			Me.mainToolbar.SuspendLayout()
			Me.SuspendLayout()
			' 
			' openFileDialog1
			' 
			Me.openFileDialog1.DefaultExt = "xls"
			Me.openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|Excel 97/2003|*.xls|Excel 2007|*.xlsx;*.xlsm|All " & "files|*.*"
			Me.openFileDialog1.Title = "Open an Excel File"
			' 
			' printDialog1
			' 
			Me.printDialog1.AllowSomePages = True
			Me.printDialog1.Document = Me.flexCelPrintDocument1
			' 
			' flexCelPrintDocument1
			' 
			Me.flexCelPrintDocument1.AllVisibleSheets = False
			Me.flexCelPrintDocument1.ResetPageNumberOnEachSheet = False
			Me.flexCelPrintDocument1.Workbook = Nothing
'			Me.flexCelPrintDocument1.GetPrinterHardMargins += New FlexCel.Render.PrintHardMarginsEventHandler(Me.flexCelPrintDocument1_GetPrinterHardMargins)
'			Me.flexCelPrintDocument1.BeforePrintPage += New System.Drawing.Printing.PrintPageEventHandler(Me.flexCelPrintDocument1_BeforePrintPage)
'			Me.flexCelPrintDocument1.PrintPage += New System.Drawing.Printing.PrintPageEventHandler(Me.flexCelPrintDocument1_PrintPage)
			' 
			' printPreviewDialog1
			' 
			Me.printPreviewDialog1.AutoScrollMargin = New System.Drawing.Size(0, 0)
			Me.printPreviewDialog1.AutoScrollMinSize = New System.Drawing.Size(0, 0)
			Me.printPreviewDialog1.ClientSize = New System.Drawing.Size(400, 300)
			Me.printPreviewDialog1.Document = Me.flexCelPrintDocument1
			Me.printPreviewDialog1.Enabled = True
			Me.printPreviewDialog1.Icon = (CType(resources.GetObject("printPreviewDialog1.Icon"), System.Drawing.Icon))
			Me.printPreviewDialog1.Name = "printPreviewDialog1"
			Me.printPreviewDialog1.Visible = False
			' 
			' panel1
			' 
			Me.panel1.BackColor = System.Drawing.Color.White
			Me.panel1.Controls.Add(Me.cbResetPageNumber)
			Me.panel1.Controls.Add(Me.panel4)
			Me.panel1.Controls.Add(Me.cbAllSheets)
			Me.panel1.Controls.Add(Me.label19)
			Me.panel1.Controls.Add(Me.cbInterpolation)
			Me.panel1.Controls.Add(Me.chHeadings)
			Me.panel1.Controls.Add(Me.cbConfidential)
			Me.panel1.Controls.Add(Me.label18)
			Me.panel1.Controls.Add(Me.cbSheet)
			Me.panel1.Controls.Add(Me.panel3)
			Me.panel1.Controls.Add(Me.Landscape)
			Me.panel1.Controls.Add(Me.label11)
			Me.panel1.Controls.Add(Me.edf)
			Me.panel1.Controls.Add(Me.labelb)
			Me.panel1.Controls.Add(Me.edb)
			Me.panel1.Controls.Add(Me.label9)
			Me.panel1.Controls.Add(Me.edr)
			Me.panel1.Controls.Add(Me.label8)
			Me.panel1.Controls.Add(Me.edt)
			Me.panel1.Controls.Add(Me.label7)
			Me.panel1.Controls.Add(Me.edl)
			Me.panel1.Controls.Add(Me.label4)
			Me.panel1.Controls.Add(Me.edZoom)
			Me.panel1.Controls.Add(Me.chFitIn)
			Me.panel1.Controls.Add(Me.chPrintLeft)
			Me.panel1.Controls.Add(Me.label6)
			Me.panel1.Controls.Add(Me.label5)
			Me.panel1.Controls.Add(Me.edVPages)
			Me.panel1.Controls.Add(Me.edHPages)
			Me.panel1.Controls.Add(Me.edFooter)
			Me.panel1.Controls.Add(Me.label3)
			Me.panel1.Controls.Add(Me.edHeader)
			Me.panel1.Controls.Add(Me.label2)
			Me.panel1.Controls.Add(Me.edFileName)
			Me.panel1.Controls.Add(Me.chFormulaText)
			Me.panel1.Controls.Add(Me.chGridLines)
			Me.panel1.Controls.Add(Me.chAntiAlias)
			Me.panel1.Controls.Add(Me.label1)
			Me.panel1.Controls.Add(Me.label10)
			Me.panel1.Controls.Add(Me.edh)
			Me.panel1.Dock = System.Windows.Forms.DockStyle.Fill
			Me.panel1.Location = New System.Drawing.Point(0, 38)
			Me.panel1.Name = "panel1"
			Me.panel1.Size = New System.Drawing.Size(768, 479)
			Me.panel1.TabIndex = 3
			' 
			' cbResetPageNumber
			' 
			Me.cbResetPageNumber.Enabled = False
			Me.cbResetPageNumber.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.cbResetPageNumber.Location = New System.Drawing.Point(528, 48)
			Me.cbResetPageNumber.Name = "cbResetPageNumber"
			Me.cbResetPageNumber.Size = New System.Drawing.Size(216, 16)
			Me.cbResetPageNumber.TabIndex = 39
			Me.cbResetPageNumber.Text = "Reset Page number on each sheet."
			' 
			' panel4
			' 
			Me.panel4.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.panel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.panel4.Location = New System.Drawing.Point(16, 72)
			Me.panel4.Name = "panel4"
			Me.panel4.Size = New System.Drawing.Size(736, 3)
			Me.panel4.TabIndex = 38
			' 
			' cbAllSheets
			' 
			Me.cbAllSheets.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.cbAllSheets.Location = New System.Drawing.Point(32, 48)
			Me.cbAllSheets.Name = "cbAllSheets"
			Me.cbAllSheets.Size = New System.Drawing.Size(104, 16)
			Me.cbAllSheets.TabIndex = 37
			Me.cbAllSheets.Text = "All Sheets"
'			Me.cbAllSheets.CheckedChanged += New System.EventHandler(Me.cbAllSheets_CheckedChanged)
			' 
			' label19
			' 
			Me.label19.Location = New System.Drawing.Point(392, 80)
			Me.label19.Name = "label19"
			Me.label19.Size = New System.Drawing.Size(160, 40)
			Me.label19.TabIndex = 36
			Me.label19.Text = "Interpolation mode for images: Sometimes a lower mode might give crisper results." & ""
			' 
			' cbInterpolation
			' 
			Me.cbInterpolation.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
			Me.cbInterpolation.Items.AddRange(New Object() { "Bicubic", "Bilinear", "Default", "High", "HighQualityBicubic", "HighQualityBilinear ", "Low", "NearestNeighbor"})
			Me.cbInterpolation.Location = New System.Drawing.Point(560, 88)
			Me.cbInterpolation.Name = "cbInterpolation"
			Me.cbInterpolation.Size = New System.Drawing.Size(152, 21)
			Me.cbInterpolation.TabIndex = 35
			' 
			' chHeadings
			' 
			Me.chHeadings.Location = New System.Drawing.Point(176, 136)
			Me.chHeadings.Name = "chHeadings"
			Me.chHeadings.Size = New System.Drawing.Size(128, 24)
			Me.chHeadings.TabIndex = 34
			Me.chHeadings.Text = "Print Headings"
			' 
			' cbConfidential
			' 
			Me.cbConfidential.Location = New System.Drawing.Point(56, 112)
			Me.cbConfidential.Name = "cbConfidential"
			Me.cbConfidential.Size = New System.Drawing.Size(232, 16)
			Me.cbConfidential.TabIndex = 33
			Me.cbConfidential.Text = "Print ""Confidential"" on each page"
			' 
			' label18
			' 
			Me.label18.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label18.Location = New System.Drawing.Point(168, 48)
			Me.label18.Name = "label18"
			Me.label18.Size = New System.Drawing.Size(88, 16)
			Me.label18.TabIndex = 32
			Me.label18.Text = "Sheet to print:"
			' 
			' cbSheet
			' 
			Me.cbSheet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
			Me.cbSheet.Location = New System.Drawing.Point(256, 43)
			Me.cbSheet.Name = "cbSheet"
			Me.cbSheet.Size = New System.Drawing.Size(160, 21)
			Me.cbSheet.TabIndex = 31
'			Me.cbSheet.SelectedIndexChanged += New System.EventHandler(Me.cbSheet_SelectedIndexChanged)
			' 
			' panel3
			' 
			Me.panel3.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.panel3.BackColor = System.Drawing.Color.FromArgb((CInt((CByte(255)))), (CInt((CByte(255)))), (CInt((CByte(192)))))
			Me.panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.panel3.Controls.Add(Me.edBottom)
			Me.panel3.Controls.Add(Me.label17)
			Me.panel3.Controls.Add(Me.edRight)
			Me.panel3.Controls.Add(Me.label16)
			Me.panel3.Controls.Add(Me.edLeft)
			Me.panel3.Controls.Add(Me.label15)
			Me.panel3.Controls.Add(Me.edTop)
			Me.panel3.Controls.Add(Me.label14)
			Me.panel3.Controls.Add(Me.label13)
			Me.panel3.Controls.Add(Me.label12)
			Me.panel3.Location = New System.Drawing.Point(504, 232)
			Me.panel3.Name = "panel3"
			Me.panel3.Size = New System.Drawing.Size(216, 224)
			Me.panel3.TabIndex = 30
			' 
			' edBottom
			' 
			Me.edBottom.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.edBottom.Location = New System.Drawing.Point(80, 136)
			Me.edBottom.Name = "edBottom"
			Me.edBottom.Size = New System.Drawing.Size(48, 20)
			Me.edBottom.TabIndex = 26
			Me.edBottom.Text = "0"
			' 
			' label17
			' 
			Me.label17.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label17.Location = New System.Drawing.Point(16, 160)
			Me.label17.Name = "label17"
			Me.label17.Size = New System.Drawing.Size(56, 16)
			Me.label17.TabIndex = 25
			Me.label17.Text = "Last Col:"
			' 
			' edRight
			' 
			Me.edRight.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.edRight.Location = New System.Drawing.Point(80, 160)
			Me.edRight.Name = "edRight"
			Me.edRight.Size = New System.Drawing.Size(48, 20)
			Me.edRight.TabIndex = 24
			Me.edRight.Text = "0"
			' 
			' label16
			' 
			Me.label16.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label16.Location = New System.Drawing.Point(16, 136)
			Me.label16.Name = "label16"
			Me.label16.Size = New System.Drawing.Size(85, 16)
			Me.label16.TabIndex = 23
			Me.label16.Text = "Last Row:"
			' 
			' edLeft
			' 
			Me.edLeft.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.edLeft.Location = New System.Drawing.Point(80, 112)
			Me.edLeft.Name = "edLeft"
			Me.edLeft.Size = New System.Drawing.Size(48, 20)
			Me.edLeft.TabIndex = 22
			Me.edLeft.Text = "0"
			' 
			' label15
			' 
			Me.label15.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label15.Location = New System.Drawing.Point(16, 112)
			Me.label15.Name = "label15"
			Me.label15.Size = New System.Drawing.Size(85, 16)
			Me.label15.TabIndex = 21
			Me.label15.Text = "First Col:"
			' 
			' edTop
			' 
			Me.edTop.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.edTop.Location = New System.Drawing.Point(80, 88)
			Me.edTop.Name = "edTop"
			Me.edTop.Size = New System.Drawing.Size(48, 20)
			Me.edTop.TabIndex = 20
			Me.edTop.Text = "0"
			' 
			' label14
			' 
			Me.label14.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label14.Location = New System.Drawing.Point(8, 88)
			Me.label14.Name = "label14"
			Me.label14.Size = New System.Drawing.Size(85, 16)
			Me.label14.TabIndex = 3
			Me.label14.Text = "First Row:"
			' 
			' label13
			' 
			Me.label13.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.label13.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label13.Location = New System.Drawing.Point(8, 32)
			Me.label13.Name = "label13"
			Me.label13.Size = New System.Drawing.Size(192, 32)
			Me.label13.TabIndex = 2
			Me.label13.Text = "If one of this values is <=0 all print_range will be printed"
			' 
			' label12
			' 
			Me.label12.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label12.Location = New System.Drawing.Point(8, 16)
			Me.label12.Name = "label12"
			Me.label12.Size = New System.Drawing.Size(192, 16)
			Me.label12.TabIndex = 1
			Me.label12.Text = "Range to Print:"
			' 
			' Landscape
			' 
			Me.Landscape.Location = New System.Drawing.Point(456, 136)
			Me.Landscape.Name = "Landscape"
			Me.Landscape.Size = New System.Drawing.Size(96, 24)
			Me.Landscape.TabIndex = 29
			Me.Landscape.Text = "Landscape"
			' 
			' label11
			' 
			Me.label11.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label11.Location = New System.Drawing.Point(264, 416)
			Me.label11.Name = "label11"
			Me.label11.Size = New System.Drawing.Size(80, 16)
			Me.label11.TabIndex = 28
			Me.label11.Text = "Footer Margin"
			' 
			' edf
			' 
			Me.edf.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.edf.Location = New System.Drawing.Point(344, 416)
			Me.edf.Name = "edf"
			Me.edf.Size = New System.Drawing.Size(128, 20)
			Me.edf.TabIndex = 27
			' 
			' labelb
			' 
			Me.labelb.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.labelb.Location = New System.Drawing.Point(256, 368)
			Me.labelb.Name = "labelb"
			Me.labelb.Size = New System.Drawing.Size(88, 16)
			Me.labelb.TabIndex = 26
			Me.labelb.Text = "Bottom Margin"
			' 
			' edb
			' 
			Me.edb.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.edb.Location = New System.Drawing.Point(344, 368)
			Me.edb.Name = "edb"
			Me.edb.Size = New System.Drawing.Size(128, 20)
			Me.edb.TabIndex = 25
			' 
			' label9
			' 
			Me.label9.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label9.Location = New System.Drawing.Point(56, 368)
			Me.label9.Name = "label9"
			Me.label9.Size = New System.Drawing.Size(80, 16)
			Me.label9.TabIndex = 24
			Me.label9.Text = "Right Margin"
			' 
			' edr
			' 
			Me.edr.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.edr.Location = New System.Drawing.Point(136, 368)
			Me.edr.Name = "edr"
			Me.edr.Size = New System.Drawing.Size(112, 20)
			Me.edr.TabIndex = 23
			' 
			' label8
			' 
			Me.label8.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label8.Location = New System.Drawing.Point(264, 328)
			Me.label8.Name = "label8"
			Me.label8.Size = New System.Drawing.Size(80, 16)
			Me.label8.TabIndex = 22
			Me.label8.Text = "Top Margin"
			' 
			' edt
			' 
			Me.edt.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.edt.Location = New System.Drawing.Point(344, 328)
			Me.edt.Name = "edt"
			Me.edt.Size = New System.Drawing.Size(128, 20)
			Me.edt.TabIndex = 21
			' 
			' label7
			' 
			Me.label7.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label7.Location = New System.Drawing.Point(56, 328)
			Me.label7.Name = "label7"
			Me.label7.Size = New System.Drawing.Size(80, 16)
			Me.label7.TabIndex = 20
			Me.label7.Text = "Left Margin"
			' 
			' edl
			' 
			Me.edl.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.edl.Location = New System.Drawing.Point(136, 328)
			Me.edl.Name = "edl"
			Me.edl.Size = New System.Drawing.Size(112, 20)
			Me.edl.TabIndex = 19
			' 
			' label4
			' 
			Me.label4.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label4.Location = New System.Drawing.Point(120, 280)
			Me.label4.Name = "label4"
			Me.label4.Size = New System.Drawing.Size(56, 16)
			Me.label4.TabIndex = 18
			Me.label4.Text = "Zoom (%)"
			' 
			' edZoom
			' 
			Me.edZoom.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.edZoom.Location = New System.Drawing.Point(184, 280)
			Me.edZoom.Name = "edZoom"
			Me.edZoom.Size = New System.Drawing.Size(24, 20)
			Me.edZoom.TabIndex = 17
			' 
			' chFitIn
			' 
			Me.chFitIn.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.chFitIn.Location = New System.Drawing.Point(56, 248)
			Me.chFitIn.Name = "chFitIn"
			Me.chFitIn.Size = New System.Drawing.Size(56, 24)
			Me.chFitIn.TabIndex = 16
			Me.chFitIn.Text = "Fit in"
'			Me.chFitIn.CheckedChanged += New System.EventHandler(Me.chFitIn_CheckedChanged)
			' 
			' chPrintLeft
			' 
			Me.chPrintLeft.Location = New System.Drawing.Point(312, 136)
			Me.chPrintLeft.Name = "chPrintLeft"
			Me.chPrintLeft.Size = New System.Drawing.Size(136, 24)
			Me.chPrintLeft.TabIndex = 15
			Me.chPrintLeft.Text = "Print Left, then down."
			' 
			' label6
			' 
			Me.label6.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label6.Location = New System.Drawing.Point(256, 248)
			Me.label6.Name = "label6"
			Me.label6.Size = New System.Drawing.Size(80, 16)
			Me.label6.TabIndex = 14
			Me.label6.Text = "pages tall."
			' 
			' label5
			' 
			Me.label5.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label5.Location = New System.Drawing.Point(144, 248)
			Me.label5.Name = "label5"
			Me.label5.Size = New System.Drawing.Size(80, 16)
			Me.label5.TabIndex = 13
			Me.label5.Text = "pages wide x"
			' 
			' edVPages
			' 
			Me.edVPages.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.edVPages.Location = New System.Drawing.Point(224, 248)
			Me.edVPages.Name = "edVPages"
			Me.edVPages.ReadOnly = True
			Me.edVPages.Size = New System.Drawing.Size(24, 20)
			Me.edVPages.TabIndex = 12
			' 
			' edHPages
			' 
			Me.edHPages.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.edHPages.Location = New System.Drawing.Point(112, 248)
			Me.edHPages.Name = "edHPages"
			Me.edHPages.ReadOnly = True
			Me.edHPages.Size = New System.Drawing.Size(24, 20)
			Me.edHPages.TabIndex = 10
			' 
			' edFooter
			' 
			Me.edFooter.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.edFooter.BackColor = System.Drawing.Color.White
			Me.edFooter.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.edFooter.Location = New System.Drawing.Point(112, 200)
			Me.edFooter.Name = "edFooter"
			Me.edFooter.Size = New System.Drawing.Size(608, 20)
			Me.edFooter.TabIndex = 8
			' 
			' label3
			' 
			Me.label3.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label3.Location = New System.Drawing.Point(56, 200)
			Me.label3.Name = "label3"
			Me.label3.Size = New System.Drawing.Size(56, 16)
			Me.label3.TabIndex = 7
			Me.label3.Text = "Footer:"
			' 
			' edHeader
			' 
			Me.edHeader.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.edHeader.BackColor = System.Drawing.Color.White
			Me.edHeader.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.edHeader.Location = New System.Drawing.Point(112, 176)
			Me.edHeader.Name = "edHeader"
			Me.edHeader.Size = New System.Drawing.Size(608, 20)
			Me.edHeader.TabIndex = 6
			' 
			' label2
			' 
			Me.label2.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label2.Location = New System.Drawing.Point(56, 176)
			Me.label2.Name = "label2"
			Me.label2.Size = New System.Drawing.Size(56, 16)
			Me.label2.TabIndex = 5
			Me.label2.Text = "Header:"
			' 
			' edFileName
			' 
			Me.edFileName.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.edFileName.BackColor = System.Drawing.Color.White
			Me.edFileName.BorderStyle = System.Windows.Forms.BorderStyle.None
			Me.edFileName.Location = New System.Drawing.Point(112, 16)
			Me.edFileName.Name = "edFileName"
			Me.edFileName.ReadOnly = True
			Me.edFileName.Size = New System.Drawing.Size(632, 13)
			Me.edFileName.TabIndex = 4
			Me.edFileName.Text = "No file selected"
			' 
			' chFormulaText
			' 
			Me.chFormulaText.Location = New System.Drawing.Point(576, 136)
			Me.chFormulaText.Name = "chFormulaText"
			Me.chFormulaText.Size = New System.Drawing.Size(136, 24)
			Me.chFormulaText.TabIndex = 3
			Me.chFormulaText.Text = "Print Formula Text"
			' 
			' chGridLines
			' 
			Me.chGridLines.Location = New System.Drawing.Point(56, 136)
			Me.chGridLines.Name = "chGridLines"
			Me.chGridLines.Size = New System.Drawing.Size(104, 24)
			Me.chGridLines.TabIndex = 2
			Me.chGridLines.Text = "Print Grid Lines"
			' 
			' chAntiAlias
			' 
			Me.chAntiAlias.Location = New System.Drawing.Point(56, 88)
			Me.chAntiAlias.Name = "chAntiAlias"
			Me.chAntiAlias.Size = New System.Drawing.Size(152, 16)
			Me.chAntiAlias.TabIndex = 1
			Me.chAntiAlias.Text = "Antialias Text"
			' 
			' label1
			' 
			Me.label1.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label1.Location = New System.Drawing.Point(24, 16)
			Me.label1.Name = "label1"
			Me.label1.Size = New System.Drawing.Size(80, 16)
			Me.label1.TabIndex = 0
			Me.label1.Text = "File to print:"
			' 
			' label10
			' 
			Me.label10.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label10.Location = New System.Drawing.Point(48, 416)
			Me.label10.Name = "label10"
			Me.label10.Size = New System.Drawing.Size(88, 16)
			Me.label10.TabIndex = 22
			Me.label10.Text = "Header Margin"
			' 
			' edh
			' 
			Me.edh.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.edh.Location = New System.Drawing.Point(136, 416)
			Me.edh.Name = "edh"
			Me.edh.Size = New System.Drawing.Size(112, 20)
			Me.edh.TabIndex = 21
			' 
			' exportImageDialog
			' 
			Me.exportImageDialog.DefaultExt = "png"
			Me.exportImageDialog.Filter = "Png files|*.png|Jpg files|*.jpg"
			Me.exportImageDialog.Title = "Save image as..."
			' 
			' exportTiffDialog
			' 
			Me.exportTiffDialog.DefaultExt = "tif"
			Me.exportTiffDialog.Filter = "TIFF Files|*.tif"
			Me.exportTiffDialog.Title = "Save image as multi page tiff..."
			' 
			' mainToolbar
			' 
			Me.mainToolbar.Items.AddRange(New System.Windows.Forms.ToolStripItem() { Me.btnOpenFile, Me.toolStripSeparator1, Me.btnSetup, Me.btnPreview, Me.btnPrint, Me.toolStripSeparator2, Me.btnExportAsImages, Me.btnExit})
			Me.mainToolbar.Location = New System.Drawing.Point(0, 0)
			Me.mainToolbar.Name = "mainToolbar"
			Me.mainToolbar.Size = New System.Drawing.Size(768, 38)
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
'			Me.btnOpenFile.Click += New System.EventHandler(Me.openFile_Click)
			' 
			' toolStripSeparator1
			' 
			Me.toolStripSeparator1.Name = "toolStripSeparator1"
			Me.toolStripSeparator1.Size = New System.Drawing.Size(6, 46)
			' 
			' btnSetup
			' 
			Me.btnSetup.Image = (CType(resources.GetObject("btnSetup.Image"), System.Drawing.Image))
			Me.btnSetup.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnSetup.Name = "btnSetup"
			Me.btnSetup.Size = New System.Drawing.Size(69, 35)
			Me.btnSetup.Text = "Print &Setup"
			Me.btnSetup.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
'			Me.btnSetup.Click += New System.EventHandler(Me.setup_Click)
			' 
			' btnPreview
			' 
			Me.btnPreview.Image = (CType(resources.GetObject("btnPreview.Image"), System.Drawing.Image))
			Me.btnPreview.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnPreview.Name = "btnPreview"
			Me.btnPreview.Size = New System.Drawing.Size(80, 35)
			Me.btnPreview.Text = "Print Pre&view"
			Me.btnPreview.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
'			Me.btnPreview.Click += New System.EventHandler(Me.preview_Click)
			' 
			' btnPrint
			' 
			Me.btnPrint.Image = (CType(resources.GetObject("btnPrint.Image"), System.Drawing.Image))
			Me.btnPrint.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnPrint.Name = "btnPrint"
			Me.btnPrint.Size = New System.Drawing.Size(36, 35)
			Me.btnPrint.Text = "&Print"
			Me.btnPrint.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
'			Me.btnPrint.Click += New System.EventHandler(Me.print_Click)
			' 
			' toolStripSeparator2
			' 
			Me.toolStripSeparator2.Name = "toolStripSeparator2"
			Me.toolStripSeparator2.Size = New System.Drawing.Size(6, 46)
			' 
			' btnExportAsImages
			' 
			Me.btnExportAsImages.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() { Me.usingFlexCelImgExportAllPagesrecommendedWayToolStripMenuItem, Me.usingFlexCelImgExport1PagerecommendedWayToolStripMenuItem, Me.multiPageTIFFUsingFlexCelImgExportToolStripMenuItem})
			Me.btnExportAsImages.Image = (CType(resources.GetObject("btnExportAsImages.Image"), System.Drawing.Image))
			Me.btnExportAsImages.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnExportAsImages.Name = "btnExportAsImages"
			Me.btnExportAsImages.Size = New System.Drawing.Size(108, 35)
			Me.btnExportAsImages.Text = "Export as &Images"
			Me.btnExportAsImages.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
			' 
			' usingFlexCelImgExportAllPagesrecommendedWayToolStripMenuItem
			' 
			Me.usingFlexCelImgExportAllPagesrecommendedWayToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() { Me.blackAndWhiteToolStripMenuItem, Me.colorsToolStripMenuItem, Me.trueColorToolStripMenuItem})
			Me.usingFlexCelImgExportAllPagesrecommendedWayToolStripMenuItem.Name = "usingFlexCelImgExportAllPagesrecommendedWayToolStripMenuItem"
			Me.usingFlexCelImgExportAllPagesrecommendedWayToolStripMenuItem.Size = New System.Drawing.Size(153, 22)
			Me.usingFlexCelImgExportAllPagesrecommendedWayToolStripMenuItem.Text = "All pages"
			' 
			' blackAndWhiteToolStripMenuItem
			' 
			Me.blackAndWhiteToolStripMenuItem.Name = "blackAndWhiteToolStripMenuItem"
			Me.blackAndWhiteToolStripMenuItem.Size = New System.Drawing.Size(161, 22)
			Me.blackAndWhiteToolStripMenuItem.Text = "Black And White"
'			Me.blackAndWhiteToolStripMenuItem.Click += New System.EventHandler(Me.ImgBlackAndWhite_Click)
			' 
			' colorsToolStripMenuItem
			' 
			Me.colorsToolStripMenuItem.Name = "colorsToolStripMenuItem"
			Me.colorsToolStripMenuItem.Size = New System.Drawing.Size(161, 22)
			Me.colorsToolStripMenuItem.Text = "256 Colors"
'			Me.colorsToolStripMenuItem.Click += New System.EventHandler(Me.Img256Colors_Click)
			' 
			' trueColorToolStripMenuItem
			' 
			Me.trueColorToolStripMenuItem.Name = "trueColorToolStripMenuItem"
			Me.trueColorToolStripMenuItem.Size = New System.Drawing.Size(161, 22)
			Me.trueColorToolStripMenuItem.Text = "True Color"
'			Me.trueColorToolStripMenuItem.Click += New System.EventHandler(Me.ImgTrueColor_Click)
			' 
			' usingFlexCelImgExport1PagerecommendedWayToolStripMenuItem
			' 
			Me.usingFlexCelImgExport1PagerecommendedWayToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() { Me.blackAndWhiteToolStripMenuItem1, Me.colorsToolStripMenuItem1, Me.trueColorToolStripMenuItem1})
			Me.usingFlexCelImgExport1PagerecommendedWayToolStripMenuItem.Name = "usingFlexCelImgExport1PagerecommendedWayToolStripMenuItem"
			Me.usingFlexCelImgExport1PagerecommendedWayToolStripMenuItem.Size = New System.Drawing.Size(153, 22)
			Me.usingFlexCelImgExport1PagerecommendedWayToolStripMenuItem.Text = "1 page"
			' 
			' blackAndWhiteToolStripMenuItem1
			' 
			Me.blackAndWhiteToolStripMenuItem1.Name = "blackAndWhiteToolStripMenuItem1"
			Me.blackAndWhiteToolStripMenuItem1.Size = New System.Drawing.Size(161, 22)
			Me.blackAndWhiteToolStripMenuItem1.Text = "Black And White"
'			Me.blackAndWhiteToolStripMenuItem1.Click += New System.EventHandler(Me.ImgBlackAndWhite2_Click)
			' 
			' colorsToolStripMenuItem1
			' 
			Me.colorsToolStripMenuItem1.Name = "colorsToolStripMenuItem1"
			Me.colorsToolStripMenuItem1.Size = New System.Drawing.Size(161, 22)
			Me.colorsToolStripMenuItem1.Text = "256 Colors"
'			Me.colorsToolStripMenuItem1.Click += New System.EventHandler(Me.Img256Colors2_Click)
			' 
			' trueColorToolStripMenuItem1
			' 
			Me.trueColorToolStripMenuItem1.Name = "trueColorToolStripMenuItem1"
			Me.trueColorToolStripMenuItem1.Size = New System.Drawing.Size(161, 22)
			Me.trueColorToolStripMenuItem1.Text = "True Color"
'			Me.trueColorToolStripMenuItem1.Click += New System.EventHandler(Me.ImgTrueColor2_Click)
			' 
			' multiPageTIFFUsingFlexCelImgExportToolStripMenuItem
			' 
			Me.multiPageTIFFUsingFlexCelImgExportToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() { Me.faxToolStripMenuItem, Me.blackAndWhiteToolStripMenuItem2, Me.colorsToolStripMenuItem2, Me.trueColorToolStripMenuItem2})
			Me.multiPageTIFFUsingFlexCelImgExportToolStripMenuItem.Name = "multiPageTIFFUsingFlexCelImgExportToolStripMenuItem"
			Me.multiPageTIFFUsingFlexCelImgExportToolStripMenuItem.Size = New System.Drawing.Size(153, 22)
			Me.multiPageTIFFUsingFlexCelImgExportToolStripMenuItem.Text = "MultiPage TIFF"
			' 
			' faxToolStripMenuItem
			' 
			Me.faxToolStripMenuItem.Name = "faxToolStripMenuItem"
			Me.faxToolStripMenuItem.Size = New System.Drawing.Size(161, 22)
			Me.faxToolStripMenuItem.Text = "Fax"
'			Me.faxToolStripMenuItem.Click += New System.EventHandler(Me.TiffFax_Click)
			' 
			' blackAndWhiteToolStripMenuItem2
			' 
			Me.blackAndWhiteToolStripMenuItem2.Name = "blackAndWhiteToolStripMenuItem2"
			Me.blackAndWhiteToolStripMenuItem2.Size = New System.Drawing.Size(161, 22)
			Me.blackAndWhiteToolStripMenuItem2.Text = "Black And White"
'			Me.blackAndWhiteToolStripMenuItem2.Click += New System.EventHandler(Me.TiffBlackAndWhite_Click)
			' 
			' colorsToolStripMenuItem2
			' 
			Me.colorsToolStripMenuItem2.Name = "colorsToolStripMenuItem2"
			Me.colorsToolStripMenuItem2.Size = New System.Drawing.Size(161, 22)
			Me.colorsToolStripMenuItem2.Text = "256 Colors"
'			Me.colorsToolStripMenuItem2.Click += New System.EventHandler(Me.Tiff256Colors_Click)
			' 
			' trueColorToolStripMenuItem2
			' 
			Me.trueColorToolStripMenuItem2.Name = "trueColorToolStripMenuItem2"
			Me.trueColorToolStripMenuItem2.Size = New System.Drawing.Size(161, 22)
			Me.trueColorToolStripMenuItem2.Text = "True Color"
'			Me.trueColorToolStripMenuItem2.Click += New System.EventHandler(Me.TiffTrueColor_Click)
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
			Me.ClientSize = New System.Drawing.Size(768, 517)
			Me.Controls.Add(Me.panel1)
			Me.Controls.Add(Me.mainToolbar)
			Me.Name = "mainForm"
			Me.Text = "Print and preview a file"
			Me.panel1.ResumeLayout(False)
			Me.panel1.PerformLayout()
			Me.panel3.ResumeLayout(False)
			Me.panel3.PerformLayout()
			Me.mainToolbar.ResumeLayout(False)
			Me.mainToolbar.PerformLayout()
			Me.ResumeLayout(False)
			Me.PerformLayout()

		End Sub
		#End Region

		Private mainToolbar As ToolStrip
		Private WithEvents btnOpenFile As ToolStripButton
		Private toolStripSeparator1 As ToolStripSeparator
		Private WithEvents btnSetup As ToolStripButton
		Private WithEvents btnPreview As ToolStripButton
		Private WithEvents btnPrint As ToolStripButton
		Private toolStripSeparator2 As ToolStripSeparator
		Private WithEvents btnExit As ToolStripButton
		Private btnExportAsImages As ToolStripDropDownButton
		Private usingFlexCelImgExportAllPagesrecommendedWayToolStripMenuItem As ToolStripMenuItem
		Private WithEvents blackAndWhiteToolStripMenuItem As ToolStripMenuItem
		Private WithEvents colorsToolStripMenuItem As ToolStripMenuItem
		Private WithEvents trueColorToolStripMenuItem As ToolStripMenuItem
		Private usingFlexCelImgExport1PagerecommendedWayToolStripMenuItem As ToolStripMenuItem
		Private WithEvents blackAndWhiteToolStripMenuItem1 As ToolStripMenuItem
		Private WithEvents colorsToolStripMenuItem1 As ToolStripMenuItem
		Private WithEvents trueColorToolStripMenuItem1 As ToolStripMenuItem
		Private multiPageTIFFUsingFlexCelImgExportToolStripMenuItem As ToolStripMenuItem
		Private WithEvents faxToolStripMenuItem As ToolStripMenuItem
		Private WithEvents blackAndWhiteToolStripMenuItem2 As ToolStripMenuItem
		Private WithEvents colorsToolStripMenuItem2 As ToolStripMenuItem
		Private WithEvents trueColorToolStripMenuItem2 As ToolStripMenuItem
	End Class
End Namespace

