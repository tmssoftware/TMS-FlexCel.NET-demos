Imports System.Collections
Imports System.ComponentModel
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Render
Imports System.IO
Imports System.Text
Namespace ExportHTML
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
		Private panel4 As System.Windows.Forms.Panel
		Private label4 As System.Windows.Forms.Label
		Private cbOutlook2007 As System.Windows.Forms.CheckBox
		Private panel5 As System.Windows.Forms.Panel
		Private label5 As System.Windows.Forms.Label
		Private panel6 As System.Windows.Forms.Panel
		Private label6 As System.Windows.Forms.Label
		Private checkBox4 As System.Windows.Forms.CheckBox
		Private cbIe6Png As System.Windows.Forms.CheckBox
		Private cbComments As System.Windows.Forms.CheckBox
		Private cbHyperlinks As System.Windows.Forms.CheckBox
		Private cbImages As System.Windows.Forms.CheckBox
		Private panel7 As System.Windows.Forms.Panel
		Private label2 As System.Windows.Forms.Label
		Private WithEvents cbExportObject As System.Windows.Forms.ComboBox
		Private lblSheetToExport As System.Windows.Forms.Label
		Private cbSheet As System.Windows.Forms.ComboBox
		Private panel9 As System.Windows.Forms.Panel
		Private label3 As System.Windows.Forms.Label
		Private cbTop As System.Windows.Forms.CheckBox
		Private cbLeft As System.Windows.Forms.CheckBox
		Private cbRight As System.Windows.Forms.CheckBox
		Private cbBottom As System.Windows.Forms.CheckBox
		Private panel10 As System.Windows.Forms.Panel
		Private label7 As System.Windows.Forms.Label
		Private edSheetSeparator As System.Windows.Forms.TextBox
		Private panel11 As System.Windows.Forms.Panel
		Private label8 As System.Windows.Forms.Label
		Private panel12 As System.Windows.Forms.Panel
		Private WithEvents cbCss As System.Windows.Forms.CheckBox
		Private edCss As System.Windows.Forms.TextBox
		Private edImages As System.Windows.Forms.TextBox
		Private cbHtmlVersion As System.Windows.Forms.ComboBox
		Private cbFileFormat As System.Windows.Forms.ComboBox
		Private label9 As System.Windows.Forms.Label
		Private panel13 As System.Windows.Forms.Panel
		Private edBodyStart As System.Windows.Forms.TextBox
		Private label10 As System.Windows.Forms.Label
		Private cbReplaceFonts As System.Windows.Forms.CheckBox
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
			Me.cbReplaceFonts = New System.Windows.Forms.CheckBox()
			Me.panel13 = New System.Windows.Forms.Panel()
			Me.edBodyStart = New System.Windows.Forms.TextBox()
			Me.label10 = New System.Windows.Forms.Label()
			Me.panel12 = New System.Windows.Forms.Panel()
			Me.cbCss = New System.Windows.Forms.CheckBox()
			Me.edCss = New System.Windows.Forms.TextBox()
			Me.panel11 = New System.Windows.Forms.Panel()
			Me.edImages = New System.Windows.Forms.TextBox()
			Me.label8 = New System.Windows.Forms.Label()
			Me.panel10 = New System.Windows.Forms.Panel()
			Me.edSheetSeparator = New System.Windows.Forms.TextBox()
			Me.label7 = New System.Windows.Forms.Label()
			Me.panel9 = New System.Windows.Forms.Panel()
			Me.cbBottom = New System.Windows.Forms.CheckBox()
			Me.cbRight = New System.Windows.Forms.CheckBox()
			Me.cbLeft = New System.Windows.Forms.CheckBox()
			Me.cbTop = New System.Windows.Forms.CheckBox()
			Me.label3 = New System.Windows.Forms.Label()
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
			Me.panel5 = New System.Windows.Forms.Panel()
			Me.cbIe6Png = New System.Windows.Forms.CheckBox()
			Me.label5 = New System.Windows.Forms.Label()
			Me.cbOutlook2007 = New System.Windows.Forms.CheckBox()
			Me.panel4 = New System.Windows.Forms.Panel()
			Me.cbEmbedImages = New System.Windows.Forms.CheckBox()
			Me.sbSVG = New System.Windows.Forms.CheckBox()
			Me.label9 = New System.Windows.Forms.Label()
			Me.cbFileFormat = New System.Windows.Forms.ComboBox()
			Me.cbHtmlVersion = New System.Windows.Forms.ComboBox()
			Me.label4 = New System.Windows.Forms.Label()
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
			Me.flexCelHtmlExport1 = New FlexCel.Render.FlexCelHtmlExport()
			Me.mainToolbar = New System.Windows.Forms.ToolStrip()
			Me.openFile = New System.Windows.Forms.ToolStripButton()
			Me.export = New System.Windows.Forms.ToolStripButton()
			Me.btnEmail = New System.Windows.Forms.ToolStripButton()
			Me.toolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
			Me.btnClose = New System.Windows.Forms.ToolStripButton()
			Me.panel1.SuspendLayout()
			Me.panel13.SuspendLayout()
			Me.panel12.SuspendLayout()
			Me.panel11.SuspendLayout()
			Me.panel10.SuspendLayout()
			Me.panel9.SuspendLayout()
			Me.panel7.SuspendLayout()
			Me.panel6.SuspendLayout()
			Me.panel5.SuspendLayout()
			Me.panel4.SuspendLayout()
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
			Me.panel1.Controls.Add(Me.cbReplaceFonts)
			Me.panel1.Controls.Add(Me.panel13)
			Me.panel1.Controls.Add(Me.panel12)
			Me.panel1.Controls.Add(Me.panel11)
			Me.panel1.Controls.Add(Me.panel10)
			Me.panel1.Controls.Add(Me.panel9)
			Me.panel1.Controls.Add(Me.panel7)
			Me.panel1.Controls.Add(Me.panel6)
			Me.panel1.Controls.Add(Me.panel5)
			Me.panel1.Controls.Add(Me.panel4)
			Me.panel1.Controls.Add(Me.panel3)
			Me.panel1.Controls.Add(Me.label1)
			Me.panel1.Controls.Add(Me.panel8)
			Me.panel1.Dock = System.Windows.Forms.DockStyle.Fill
			Me.panel1.Location = New System.Drawing.Point(0, 0)
			Me.panel1.Name = "panel1"
			Me.panel1.Size = New System.Drawing.Size(768, 696)
			Me.panel1.TabIndex = 3
			' 
			' cbReplaceFonts
			' 
			Me.cbReplaceFonts.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.cbReplaceFonts.Location = New System.Drawing.Point(40, 634)
			Me.cbReplaceFonts.Name = "cbReplaceFonts"
			Me.cbReplaceFonts.Size = New System.Drawing.Size(632, 24)
			Me.cbReplaceFonts.TabIndex = 50
			Me.cbReplaceFonts.Text = "Replace all fonts with Arial"
			' 
			' panel13
			' 
			Me.panel13.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.panel13.BackColor = System.Drawing.Color.FromArgb((CInt((CByte(224)))), (CInt((CByte(224)))), (CInt((CByte(224)))))
			Me.panel13.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.panel13.Controls.Add(Me.edBodyStart)
			Me.panel13.Controls.Add(Me.label10)
			Me.panel13.Location = New System.Drawing.Point(32, 562)
			Me.panel13.Name = "panel13"
			Me.panel13.Size = New System.Drawing.Size(704, 64)
			Me.panel13.TabIndex = 49
			' 
			' edBodyStart
			' 
			Me.edBodyStart.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.edBodyStart.Location = New System.Drawing.Point(16, 32)
			Me.edBodyStart.Name = "edBodyStart"
			Me.edBodyStart.Size = New System.Drawing.Size(664, 20)
			Me.edBodyStart.TabIndex = 20
			Me.edBodyStart.Text = "<h1>Created with FlexCel</h1>"
			' 
			' label10
			' 
			Me.label10.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label10.Location = New System.Drawing.Point(8, 8)
			Me.label10.Name = "label10"
			Me.label10.Size = New System.Drawing.Size(584, 16)
			Me.label10.TabIndex = 19
			Me.label10.Text = "Text to add at the beginning of the file:"
			' 
			' panel12
			' 
			Me.panel12.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.panel12.BackColor = System.Drawing.Color.FromArgb((CInt((CByte(224)))), (CInt((CByte(224)))), (CInt((CByte(224)))))
			Me.panel12.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.panel12.Controls.Add(Me.cbCss)
			Me.panel12.Controls.Add(Me.edCss)
			Me.panel12.Location = New System.Drawing.Point(32, 490)
			Me.panel12.Name = "panel12"
			Me.panel12.Size = New System.Drawing.Size(704, 64)
			Me.panel12.TabIndex = 48
			' 
			' cbCss
			' 
			Me.cbCss.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.cbCss.Location = New System.Drawing.Point(16, 8)
			Me.cbCss.Name = "cbCss"
			Me.cbCss.Size = New System.Drawing.Size(632, 24)
			Me.cbCss.TabIndex = 21
			Me.cbCss.Text = "Save CSS to an external file (path relative to where the html file is)"
'			Me.cbCss.CheckedChanged += New System.EventHandler(Me.cbCss_CheckedChanged)
			' 
			' edCss
			' 
			Me.edCss.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.edCss.Enabled = False
			Me.edCss.Location = New System.Drawing.Point(16, 32)
			Me.edCss.Name = "edCss"
			Me.edCss.Size = New System.Drawing.Size(664, 20)
			Me.edCss.TabIndex = 20
			Me.edCss.Text = "css\test.css"
			' 
			' panel11
			' 
			Me.panel11.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.panel11.BackColor = System.Drawing.Color.FromArgb((CInt((CByte(224)))), (CInt((CByte(224)))), (CInt((CByte(224)))))
			Me.panel11.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.panel11.Controls.Add(Me.edImages)
			Me.panel11.Controls.Add(Me.label8)
			Me.panel11.Location = New System.Drawing.Point(32, 418)
			Me.panel11.Name = "panel11"
			Me.panel11.Size = New System.Drawing.Size(704, 64)
			Me.panel11.TabIndex = 47
			' 
			' edImages
			' 
			Me.edImages.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.edImages.Location = New System.Drawing.Point(16, 32)
			Me.edImages.Name = "edImages"
			Me.edImages.Size = New System.Drawing.Size(664, 20)
			Me.edImages.TabIndex = 20
			Me.edImages.Text = "images"
			' 
			' label8
			' 
			Me.label8.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label8.Location = New System.Drawing.Point(8, 8)
			Me.label8.Name = "label8"
			Me.label8.Size = New System.Drawing.Size(584, 16)
			Me.label8.TabIndex = 19
			Me.label8.Text = "Relative path for images (make it empty to save the images in the same folder as " & "the html file)"
			' 
			' panel10
			' 
			Me.panel10.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.panel10.BackColor = System.Drawing.Color.FromArgb((CInt((CByte(224)))), (CInt((CByte(224)))), (CInt((CByte(224)))))
			Me.panel10.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.panel10.Controls.Add(Me.edSheetSeparator)
			Me.panel10.Controls.Add(Me.label7)
			Me.panel10.Location = New System.Drawing.Point(224, 338)
			Me.panel10.Name = "panel10"
			Me.panel10.Size = New System.Drawing.Size(512, 72)
			Me.panel10.TabIndex = 46
			' 
			' edSheetSeparator
			' 
			Me.edSheetSeparator.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.edSheetSeparator.Location = New System.Drawing.Point(16, 32)
			Me.edSheetSeparator.Name = "edSheetSeparator"
			Me.edSheetSeparator.Size = New System.Drawing.Size(472, 20)
			Me.edSheetSeparator.TabIndex = 20
			Me.edSheetSeparator.Text = "<p><hr />Sheet <#SheetName>  (<#SheetPos> of <#SheetCount>)"
			' 
			' label7
			' 
			Me.label7.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label7.Location = New System.Drawing.Point(8, 8)
			Me.label7.Name = "label7"
			Me.label7.Size = New System.Drawing.Size(480, 16)
			Me.label7.TabIndex = 19
			Me.label7.Text = "Sheet separator (When exporting all sheets in one file)"
			' 
			' panel9
			' 
			Me.panel9.BackColor = System.Drawing.Color.FromArgb((CInt((CByte(224)))), (CInt((CByte(224)))), (CInt((CByte(224)))))
			Me.panel9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.panel9.Controls.Add(Me.cbBottom)
			Me.panel9.Controls.Add(Me.cbRight)
			Me.panel9.Controls.Add(Me.cbLeft)
			Me.panel9.Controls.Add(Me.cbTop)
			Me.panel9.Controls.Add(Me.label3)
			Me.panel9.Location = New System.Drawing.Point(224, 266)
			Me.panel9.Name = "panel9"
			Me.panel9.Size = New System.Drawing.Size(288, 64)
			Me.panel9.TabIndex = 45
			' 
			' cbBottom
			' 
			Me.cbBottom.Location = New System.Drawing.Point(216, 32)
			Me.cbBottom.Name = "cbBottom"
			Me.cbBottom.Size = New System.Drawing.Size(64, 16)
			Me.cbBottom.TabIndex = 23
			Me.cbBottom.Text = "Bottom"
			' 
			' cbRight
			' 
			Me.cbRight.Location = New System.Drawing.Point(144, 32)
			Me.cbRight.Name = "cbRight"
			Me.cbRight.Size = New System.Drawing.Size(64, 16)
			Me.cbRight.TabIndex = 22
			Me.cbRight.Text = "Right"
			' 
			' cbLeft
			' 
			Me.cbLeft.Location = New System.Drawing.Point(24, 32)
			Me.cbLeft.Name = "cbLeft"
			Me.cbLeft.Size = New System.Drawing.Size(48, 16)
			Me.cbLeft.TabIndex = 21
			Me.cbLeft.Text = "Left"
			' 
			' cbTop
			' 
			Me.cbTop.Checked = True
			Me.cbTop.CheckState = System.Windows.Forms.CheckState.Checked
			Me.cbTop.Location = New System.Drawing.Point(88, 32)
			Me.cbTop.Name = "cbTop"
			Me.cbTop.Size = New System.Drawing.Size(48, 16)
			Me.cbTop.TabIndex = 20
			Me.cbTop.Text = "Top"
			' 
			' label3
			' 
			Me.label3.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label3.Location = New System.Drawing.Point(8, 8)
			Me.label3.Name = "label3"
			Me.label3.Size = New System.Drawing.Size(192, 16)
			Me.label3.TabIndex = 19
			Me.label3.Text = "Tabs:"
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
			Me.panel7.Size = New System.Drawing.Size(704, 72)
			Me.panel7.TabIndex = 44
			' 
			' cbExportObject
			' 
			Me.cbExportObject.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
			Me.cbExportObject.Items.AddRange(New Object() { "All sheets as Tabs", "All sheets as a single file", "Active Sheet:"})
			Me.cbExportObject.Location = New System.Drawing.Point(8, 32)
			Me.cbExportObject.Name = "cbExportObject"
			Me.cbExportObject.Size = New System.Drawing.Size(248, 21)
			Me.cbExportObject.TabIndex = 46
'			Me.cbExportObject.SelectedIndexChanged += New System.EventHandler(Me.cbExportObject_SelectedIndexChanged)
			' 
			' lblSheetToExport
			' 
			Me.lblSheetToExport.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.lblSheetToExport.Location = New System.Drawing.Point(304, 8)
			Me.lblSheetToExport.Name = "lblSheetToExport"
			Me.lblSheetToExport.Size = New System.Drawing.Size(96, 16)
			Me.lblSheetToExport.TabIndex = 45
			Me.lblSheetToExport.Text = "Sheet to export:"
			' 
			' cbSheet
			' 
			Me.cbSheet.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.cbSheet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
			Me.cbSheet.Location = New System.Drawing.Point(304, 32)
			Me.cbSheet.Name = "cbSheet"
			Me.cbSheet.Size = New System.Drawing.Size(360, 21)
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
			Me.panel6.BackColor = System.Drawing.Color.FromArgb((CInt((CByte(224)))), (CInt((CByte(224)))), (CInt((CByte(224)))))
			Me.panel6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.panel6.Controls.Add(Me.cbHeadersFooters)
			Me.panel6.Controls.Add(Me.cbImages)
			Me.panel6.Controls.Add(Me.cbHyperlinks)
			Me.panel6.Controls.Add(Me.cbComments)
			Me.panel6.Controls.Add(Me.label6)
			Me.panel6.Location = New System.Drawing.Point(32, 226)
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
			' panel5
			' 
			Me.panel5.BackColor = System.Drawing.Color.FromArgb((CInt((CByte(224)))), (CInt((CByte(224)))), (CInt((CByte(224)))))
			Me.panel5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.panel5.Controls.Add(Me.cbIe6Png)
			Me.panel5.Controls.Add(Me.label5)
			Me.panel5.Controls.Add(Me.cbOutlook2007)
			Me.panel5.Location = New System.Drawing.Point(32, 338)
			Me.panel5.Name = "panel5"
			Me.panel5.Size = New System.Drawing.Size(176, 72)
			Me.panel5.TabIndex = 41
			' 
			' cbIe6Png
			' 
			Me.cbIe6Png.Location = New System.Drawing.Point(16, 40)
			Me.cbIe6Png.Name = "cbIe6Png"
			Me.cbIe6Png.Size = New System.Drawing.Size(153, 24)
			Me.cbIe6Png.TabIndex = 20
			Me.cbIe6Png.Text = "Fix for IE6 support"
			' 
			' label5
			' 
			Me.label5.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label5.Location = New System.Drawing.Point(8, 8)
			Me.label5.Name = "label5"
			Me.label5.Size = New System.Drawing.Size(192, 16)
			Me.label5.TabIndex = 19
			Me.label5.Text = "Browser Fixes"
			' 
			' cbOutlook2007
			' 
			Me.cbOutlook2007.Location = New System.Drawing.Point(16, 24)
			Me.cbOutlook2007.Name = "cbOutlook2007"
			Me.cbOutlook2007.Size = New System.Drawing.Size(128, 16)
			Me.cbOutlook2007.TabIndex = 16
			Me.cbOutlook2007.Text = "Fix for Outlook2007"
			' 
			' panel4
			' 
			Me.panel4.BackColor = System.Drawing.Color.FromArgb((CInt((CByte(224)))), (CInt((CByte(224)))), (CInt((CByte(224)))))
			Me.panel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.panel4.Controls.Add(Me.cbEmbedImages)
			Me.panel4.Controls.Add(Me.sbSVG)
			Me.panel4.Controls.Add(Me.label9)
			Me.panel4.Controls.Add(Me.cbFileFormat)
			Me.panel4.Controls.Add(Me.cbHtmlVersion)
			Me.panel4.Controls.Add(Me.label4)
			Me.panel4.Location = New System.Drawing.Point(224, 130)
			Me.panel4.Name = "panel4"
			Me.panel4.Size = New System.Drawing.Size(288, 128)
			Me.panel4.TabIndex = 40
			' 
			' cbEmbedImages
			' 
			Me.cbEmbedImages.Location = New System.Drawing.Point(8, 91)
			Me.cbEmbedImages.Name = "cbEmbedImages"
			Me.cbEmbedImages.Size = New System.Drawing.Size(128, 28)
			Me.cbEmbedImages.TabIndex = 51
			Me.cbEmbedImages.Text = "Embed images"
			' 
			' sbSVG
			' 
			Me.sbSVG.Location = New System.Drawing.Point(141, 91)
			Me.sbSVG.Name = "sbSVG"
			Me.sbSVG.Size = New System.Drawing.Size(139, 28)
			Me.sbSVG.TabIndex = 50
			Me.sbSVG.Text = "Export images as SVG"
			' 
			' label9
			' 
			Me.label9.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label9.Location = New System.Drawing.Point(8, 8)
			Me.label9.Name = "label9"
			Me.label9.Size = New System.Drawing.Size(192, 16)
			Me.label9.TabIndex = 49
			Me.label9.Text = "HTML Version"
			' 
			' cbFileFormat
			' 
			Me.cbFileFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
			Me.cbFileFormat.Items.AddRange(New Object() { "HTML", "MHTML"})
			Me.cbFileFormat.Location = New System.Drawing.Point(8, 64)
			Me.cbFileFormat.Name = "cbFileFormat"
			Me.cbFileFormat.Size = New System.Drawing.Size(272, 21)
			Me.cbFileFormat.TabIndex = 48
			' 
			' cbHtmlVersion
			' 
			Me.cbHtmlVersion.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
			Me.cbHtmlVersion.Items.AddRange(New Object() { "HTML 3.2", "HTML 4.01", "XHTML 1.1", "HTML 5"})
			Me.cbHtmlVersion.Location = New System.Drawing.Point(8, 24)
			Me.cbHtmlVersion.Name = "cbHtmlVersion"
			Me.cbHtmlVersion.Size = New System.Drawing.Size(272, 21)
			Me.cbHtmlVersion.TabIndex = 47
			' 
			' label4
			' 
			Me.label4.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label4.Location = New System.Drawing.Point(5, 48)
			Me.label4.Name = "label4"
			Me.label4.Size = New System.Drawing.Size(192, 16)
			Me.label4.TabIndex = 19
			Me.label4.Text = "File Format:"
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
			Me.panel3.Location = New System.Drawing.Point(528, 130)
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
			Me.panel8.BackColor = System.Drawing.Color.FromArgb((CInt((CByte(224)))), (CInt((CByte(224)))), (CInt((CByte(224)))))
			Me.panel8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.panel8.Controls.Add(Me.chPrintHeadings)
			Me.panel8.Controls.Add(Me.label24)
			Me.panel8.Controls.Add(Me.chFormulaText)
			Me.panel8.Controls.Add(Me.chGridLines)
			Me.panel8.Location = New System.Drawing.Point(32, 130)
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
			Me.exportDialog.DefaultExt = "htm"
			Me.exportDialog.Filter = "HTML files|*.htm|MHTML files|*.mht"
			' 
			' flexCelHtmlExport1
			' 
			Me.flexCelHtmlExport1.HeadingWidth = 50R
			Me.flexCelHtmlExport1.ImageResolution = 96R
			Me.flexCelHtmlExport1.UsePrintScale = False
			Me.flexCelHtmlExport1.Workbook = Nothing
'			Me.flexCelHtmlExport1.HtmlFont += New FlexCel.Core.HtmlFontEventHandler(Me.flexCelHtmlExport1_HtmlFont)
			' 
			' mainToolbar
			' 
			Me.mainToolbar.ImageScalingSize = New System.Drawing.Size(24, 24)
			Me.mainToolbar.Items.AddRange(New System.Windows.Forms.ToolStripItem() { Me.openFile, Me.export, Me.btnEmail, Me.toolStripSeparator1, Me.btnClose})
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
			Me.export.Size = New System.Drawing.Size(118, 28)
			Me.export.Text = "Export as HTML"
'			Me.export.Click += New System.EventHandler(Me.export_Click)
			' 
			' btnEmail
			' 
			Me.btnEmail.Image = (CType(resources.GetObject("btnEmail.Image"), System.Drawing.Image))
			Me.btnEmail.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnEmail.Name = "btnEmail"
			Me.btnEmail.Size = New System.Drawing.Size(133, 28)
			Me.btnEmail.Text = "Email (as MHTML)"
'			Me.btnEmail.Click += New System.EventHandler(Me.btnEmail_Click)
			' 
			' toolStripSeparator1
			' 
			Me.toolStripSeparator1.Name = "toolStripSeparator1"
			Me.toolStripSeparator1.Size = New System.Drawing.Size(6, 31)
			' 
			' btnClose
			' 
			Me.btnClose.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
			Me.btnClose.Image = (CType(resources.GetObject("btnClose.Image"), System.Drawing.Image))
			Me.btnClose.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnClose.Name = "btnClose"
			Me.btnClose.Size = New System.Drawing.Size(53, 28)
			Me.btnClose.Text = "Exit"
'			Me.btnClose.Click += New System.EventHandler(Me.button2_Click)
			' 
			' mainForm
			' 
			Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New System.Drawing.Size(768, 696)
			Me.Controls.Add(Me.mainToolbar)
			Me.Controls.Add(Me.panel1)
			Me.Name = "mainForm"
			Me.Text = "Export an Excel file to HTML"
'			Me.Load += New System.EventHandler(Me.mainForm_Load)
			Me.panel1.ResumeLayout(False)
			Me.panel13.ResumeLayout(False)
			Me.panel13.PerformLayout()
			Me.panel12.ResumeLayout(False)
			Me.panel12.PerformLayout()
			Me.panel11.ResumeLayout(False)
			Me.panel11.PerformLayout()
			Me.panel10.ResumeLayout(False)
			Me.panel10.PerformLayout()
			Me.panel9.ResumeLayout(False)
			Me.panel7.ResumeLayout(False)
			Me.panel6.ResumeLayout(False)
			Me.panel5.ResumeLayout(False)
			Me.panel4.ResumeLayout(False)
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
		Private WithEvents btnEmail As ToolStripButton
		Private toolStripSeparator1 As ToolStripSeparator
		Private WithEvents btnClose As ToolStripButton
		Private sbSVG As CheckBox
		Private cbEmbedImages As CheckBox
	End Class
End Namespace

