Imports System.Collections
Imports System.ComponentModel
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Render
Imports System.IO
Imports System.Text
Namespace ExportSVG
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Private openFileDialog1 As System.Windows.Forms.OpenFileDialog
		Private panel1 As System.Windows.Forms.Panel
		Private label1 As System.Windows.Forms.Label
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
		Private exportDialog As System.Windows.Forms.SaveFileDialog
		Private panel8 As System.Windows.Forms.Panel
		Private chFormulaText As System.Windows.Forms.CheckBox
		Private chGridLines As System.Windows.Forms.CheckBox
		Private label24 As System.Windows.Forms.Label
		Private panel6 As System.Windows.Forms.Panel
		Private label6 As System.Windows.Forms.Label
		Private checkBox4 As System.Windows.Forms.CheckBox
		Private cbComments As System.Windows.Forms.CheckBox
		Private cbHyperlinks As System.Windows.Forms.CheckBox
		Private cbImages As System.Windows.Forms.CheckBox
		Private panel7 As System.Windows.Forms.Panel
		Private label2 As System.Windows.Forms.Label
		Private WithEvents cbExportObject As System.Windows.Forms.ComboBox
		Private lblSheetToExport As System.Windows.Forms.Label
		Private cbSheet As System.Windows.Forms.ComboBox
		Private chPrintHeadings As System.Windows.Forms.CheckBox
		Private cbHeadersFooters As System.Windows.Forms.CheckBox
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
			Me.panel7 = New System.Windows.Forms.Panel()
			Me.cbExportObject = New System.Windows.Forms.ComboBox()
			Me.lblSheetToExport = New System.Windows.Forms.Label()
			Me.cbSheet = New System.Windows.Forms.ComboBox()
			Me.label2 = New System.Windows.Forms.Label()
			Me.panel6 = New System.Windows.Forms.Panel()
			Me.cbHeadersFooters = New System.Windows.Forms.CheckBox()
			Me.cbImages = New System.Windows.Forms.CheckBox()
			Me.cbHyperlinks = New System.Windows.Forms.CheckBox()
			Me.cbComments = New System.Windows.Forms.CheckBox()
			Me.label6 = New System.Windows.Forms.Label()
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
			Me.label1 = New System.Windows.Forms.Label()
			Me.panel8 = New System.Windows.Forms.Panel()
			Me.chPrintHeadings = New System.Windows.Forms.CheckBox()
			Me.label24 = New System.Windows.Forms.Label()
			Me.chFormulaText = New System.Windows.Forms.CheckBox()
			Me.chGridLines = New System.Windows.Forms.CheckBox()
			Me.checkBox4 = New System.Windows.Forms.CheckBox()
			Me.exportDialog = New System.Windows.Forms.SaveFileDialog()
			Me.mainToolbar = New System.Windows.Forms.ToolStrip()
			Me.openFile = New System.Windows.Forms.ToolStripButton()
			Me.export = New System.Windows.Forms.ToolStripButton()
			Me.btnClose = New System.Windows.Forms.ToolStripButton()
			Me.panel1.SuspendLayout()
			Me.panel7.SuspendLayout()
			Me.panel6.SuspendLayout()
			Me.panel3.SuspendLayout()
			Me.panel8.SuspendLayout()
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
			Me.panel1.BackColor = System.Drawing.Color.White
			Me.panel1.Controls.Add(Me.panel7)
			Me.panel1.Controls.Add(Me.panel6)
			Me.panel1.Controls.Add(Me.panel3)
			Me.panel1.Controls.Add(Me.label1)
			Me.panel1.Controls.Add(Me.panel8)
			Me.panel1.Dock = System.Windows.Forms.DockStyle.Fill
			Me.panel1.Location = New System.Drawing.Point(0, 0)
			Me.panel1.Name = "panel1"
			Me.panel1.Size = New System.Drawing.Size(768, 268)
			Me.panel1.TabIndex = 3
			' 
			' panel7
			' 
			Me.panel7.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.panel7.BackColor = System.Drawing.Color.FromArgb((CInt((CByte(224)))), (CInt((CByte(224)))), (CInt((CByte(224)))))
			Me.panel7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.panel7.Controls.Add(Me.cbExportObject)
			Me.panel7.Controls.Add(Me.lblSheetToExport)
			Me.panel7.Controls.Add(Me.cbSheet)
			Me.panel7.Controls.Add(Me.label2)
			Me.panel7.Location = New System.Drawing.Point(32, 52)
			Me.panel7.Name = "panel7"
			Me.panel7.Size = New System.Drawing.Size(328, 200)
			Me.panel7.TabIndex = 44
			' 
			' cbExportObject
			' 
			Me.cbExportObject.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.cbExportObject.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
			Me.cbExportObject.Items.AddRange(New Object() { "All sheets", "Active Sheet:"})
			Me.cbExportObject.Location = New System.Drawing.Point(8, 32)
			Me.cbExportObject.Name = "cbExportObject"
			Me.cbExportObject.Size = New System.Drawing.Size(293, 21)
			Me.cbExportObject.TabIndex = 46
'			Me.cbExportObject.SelectedIndexChanged += New System.EventHandler(Me.cbExportObject_SelectedIndexChanged)
			' 
			' lblSheetToExport
			' 
			Me.lblSheetToExport.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.lblSheetToExport.Location = New System.Drawing.Point(8, 66)
			Me.lblSheetToExport.Name = "lblSheetToExport"
			Me.lblSheetToExport.Size = New System.Drawing.Size(96, 16)
			Me.lblSheetToExport.TabIndex = 45
			Me.lblSheetToExport.Text = "Sheet to export:"
			' 
			' cbSheet
			' 
			Me.cbSheet.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.cbSheet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
			Me.cbSheet.Location = New System.Drawing.Point(8, 82)
			Me.cbSheet.Name = "cbSheet"
			Me.cbSheet.Size = New System.Drawing.Size(294, 21)
			Me.cbSheet.TabIndex = 44
			' 
			' label2
			' 
			Me.label2.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label2.Location = New System.Drawing.Point(8, 8)
			Me.label2.Name = "label2"
			Me.label2.Size = New System.Drawing.Size(192, 16)
			Me.label2.TabIndex = 19
			Me.label2.Text = "What to Export:"
			' 
			' panel6
			' 
			Me.panel6.Anchor = (CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.panel6.BackColor = System.Drawing.Color.FromArgb((CInt((CByte(224)))), (CInt((CByte(224)))), (CInt((CByte(224)))))
			Me.panel6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.panel6.Controls.Add(Me.cbHeadersFooters)
			Me.panel6.Controls.Add(Me.cbImages)
			Me.panel6.Controls.Add(Me.cbHyperlinks)
			Me.panel6.Controls.Add(Me.cbComments)
			Me.panel6.Controls.Add(Me.label6)
			Me.panel6.Location = New System.Drawing.Point(366, 148)
			Me.panel6.Name = "panel6"
			Me.panel6.Size = New System.Drawing.Size(176, 104)
			Me.panel6.TabIndex = 42
			' 
			' cbHeadersFooters
			' 
			Me.cbHeadersFooters.Location = New System.Drawing.Point(96, 40)
			Me.cbHeadersFooters.Name = "cbHeadersFooters"
			Me.cbHeadersFooters.Size = New System.Drawing.Size(72, 44)
			Me.cbHeadersFooters.TabIndex = 23
			Me.cbHeadersFooters.Text = "Headers / Footers"
			' 
			' cbImages
			' 
			Me.cbImages.Checked = True
			Me.cbImages.CheckState = System.Windows.Forms.CheckState.Checked
			Me.cbImages.Location = New System.Drawing.Point(16, 32)
			Me.cbImages.Name = "cbImages"
			Me.cbImages.Size = New System.Drawing.Size(72, 24)
			Me.cbImages.TabIndex = 22
			Me.cbImages.Text = "Images"
			' 
			' cbHyperlinks
			' 
			Me.cbHyperlinks.Checked = True
			Me.cbHyperlinks.CheckState = System.Windows.Forms.CheckState.Checked
			Me.cbHyperlinks.Location = New System.Drawing.Point(16, 80)
			Me.cbHyperlinks.Name = "cbHyperlinks"
			Me.cbHyperlinks.Size = New System.Drawing.Size(80, 24)
			Me.cbHyperlinks.TabIndex = 21
			Me.cbHyperlinks.Text = "HyperLinks"
			' 
			' cbComments
			' 
			Me.cbComments.Checked = True
			Me.cbComments.CheckState = System.Windows.Forms.CheckState.Checked
			Me.cbComments.Location = New System.Drawing.Point(16, 56)
			Me.cbComments.Name = "cbComments"
			Me.cbComments.Size = New System.Drawing.Size(80, 24)
			Me.cbComments.TabIndex = 20
			Me.cbComments.Text = "Comments"
			' 
			' label6
			' 
			Me.label6.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label6.Location = New System.Drawing.Point(8, 16)
			Me.label6.Name = "label6"
			Me.label6.Size = New System.Drawing.Size(192, 16)
			Me.label6.TabIndex = 19
			Me.label6.Text = "Objects to Export:"
			' 
			' panel3
			' 
			Me.panel3.Anchor = (CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
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
			Me.panel3.Location = New System.Drawing.Point(548, 52)
			Me.panel3.Name = "panel3"
			Me.panel3.Size = New System.Drawing.Size(208, 200)
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
			Me.label14.Location = New System.Drawing.Point(16, 88)
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
			Me.label13.Size = New System.Drawing.Size(184, 32)
			Me.label13.TabIndex = 2
			Me.label13.Text = "If any value is <=0 all print_range will be printed"
			' 
			' label12
			' 
			Me.label12.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label12.Location = New System.Drawing.Point(8, 16)
			Me.label12.Name = "label12"
			Me.label12.Size = New System.Drawing.Size(192, 16)
			Me.label12.TabIndex = 1
			Me.label12.Text = "Range to Export:"
			' 
			' label1
			' 
			Me.label1.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label1.Location = New System.Drawing.Point(40, 16)
			Me.label1.Name = "label1"
			Me.label1.Size = New System.Drawing.Size(80, 16)
			Me.label1.TabIndex = 0
			Me.label1.Text = "File to export:"
			' 
			' panel8
			' 
			Me.panel8.Anchor = (CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.panel8.BackColor = System.Drawing.Color.FromArgb((CInt((CByte(224)))), (CInt((CByte(224)))), (CInt((CByte(224)))))
			Me.panel8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.panel8.Controls.Add(Me.chPrintHeadings)
			Me.panel8.Controls.Add(Me.label24)
			Me.panel8.Controls.Add(Me.chFormulaText)
			Me.panel8.Controls.Add(Me.chGridLines)
			Me.panel8.Location = New System.Drawing.Point(366, 52)
			Me.panel8.Name = "panel8"
			Me.panel8.Size = New System.Drawing.Size(176, 88)
			Me.panel8.TabIndex = 37
			' 
			' chPrintHeadings
			' 
			Me.chPrintHeadings.Location = New System.Drawing.Point(16, 44)
			Me.chPrintHeadings.Name = "chPrintHeadings"
			Me.chPrintHeadings.Size = New System.Drawing.Size(144, 16)
			Me.chPrintHeadings.TabIndex = 20
			Me.chPrintHeadings.Text = "Print Headings"
			' 
			' label24
			' 
			Me.label24.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label24.Location = New System.Drawing.Point(8, 8)
			Me.label24.Name = "label24"
			Me.label24.Size = New System.Drawing.Size(192, 16)
			Me.label24.TabIndex = 19
			Me.label24.Text = "Export Options:"
			' 
			' chFormulaText
			' 
			Me.chFormulaText.Location = New System.Drawing.Point(16, 64)
			Me.chFormulaText.Name = "chFormulaText"
			Me.chFormulaText.Size = New System.Drawing.Size(136, 16)
			Me.chFormulaText.TabIndex = 17
			Me.chFormulaText.Text = "Print Formula Text"
			' 
			' chGridLines
			' 
			Me.chGridLines.Location = New System.Drawing.Point(16, 24)
			Me.chGridLines.Name = "chGridLines"
			Me.chGridLines.Size = New System.Drawing.Size(128, 16)
			Me.chGridLines.TabIndex = 16
			Me.chGridLines.Text = "Print Grid Lines"
			' 
			' checkBox4
			' 
			Me.checkBox4.Location = New System.Drawing.Point(0, 0)
			Me.checkBox4.Name = "checkBox4"
			Me.checkBox4.Size = New System.Drawing.Size(104, 24)
			Me.checkBox4.TabIndex = 0
			' 
			' exportDialog
			' 
			Me.exportDialog.DefaultExt = "svg"
			Me.exportDialog.Filter = "SVG Files|*.svg"
			Me.exportDialog.Title = "Files will be saved as Filename_sheetname_pagenumber.svg"
			' 
			' mainToolbar
			' 
			Me.mainToolbar.ImageScalingSize = New System.Drawing.Size(24, 24)
			Me.mainToolbar.Items.AddRange(New System.Windows.Forms.ToolStripItem() { Me.openFile, Me.export, Me.btnClose})
			Me.mainToolbar.Location = New System.Drawing.Point(0, 0)
			Me.mainToolbar.Name = "mainToolbar"
			Me.mainToolbar.Size = New System.Drawing.Size(768, 31)
			Me.mainToolbar.TabIndex = 8
			' 
			' openFile
			' 
			Me.openFile.Image = (CType(resources.GetObject("openFile.Image"), System.Drawing.Image))
			Me.openFile.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.openFile.Name = "openFile"
			Me.openFile.Size = New System.Drawing.Size(85, 28)
			Me.openFile.Text = "Open File"
'			Me.openFile.Click += New System.EventHandler(Me.openFile_Click)
			' 
			' export
			' 
			Me.export.Image = (CType(resources.GetObject("export.Image"), System.Drawing.Image))
			Me.export.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.export.Name = "export"
			Me.export.Size = New System.Drawing.Size(106, 28)
			Me.export.Text = "Export as SVG"
'			Me.export.Click += New System.EventHandler(Me.export_Click)
			' 
			' btnClose
			' 
			Me.btnClose.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
			Me.btnClose.Image = (CType(resources.GetObject("btnClose.Image"), System.Drawing.Image))
			Me.btnClose.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnClose.Name = "btnClose"
			Me.btnClose.Size = New System.Drawing.Size(53, 28)
			Me.btnClose.Text = "Exit"
'			Me.btnClose.Click += New System.EventHandler(Me.btnClose_Click)
			' 
			' mainForm
			' 
			Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New System.Drawing.Size(768, 268)
			Me.Controls.Add(Me.mainToolbar)
			Me.Controls.Add(Me.panel1)
			Me.Name = "mainForm"
			Me.Text = "Export an Excel file to SVG"
'			Me.Load += New System.EventHandler(Me.mainForm_Load)
			Me.panel1.ResumeLayout(False)
			Me.panel7.ResumeLayout(False)
			Me.panel6.ResumeLayout(False)
			Me.panel3.ResumeLayout(False)
			Me.panel3.PerformLayout()
			Me.panel8.ResumeLayout(False)
			Me.mainToolbar.ResumeLayout(False)
			Me.mainToolbar.PerformLayout()
			Me.ResumeLayout(False)
			Me.PerformLayout()

		End Sub
		#End Region

		Private mainToolbar As ToolStrip
		Private WithEvents openFile As ToolStripButton
		Private WithEvents export As ToolStripButton
		Private WithEvents btnClose As ToolStripButton
	End Class
End Namespace

