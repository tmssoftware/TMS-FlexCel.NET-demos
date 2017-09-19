Imports System.Collections
Imports System.ComponentModel
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports System.IO
Imports System.Reflection
Imports System.Drawing.Drawing2D
Imports FlexCel.Pdf
Imports System.Runtime.InteropServices
Namespace ExportPdf
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Private openFileDialog1 As System.Windows.Forms.OpenFileDialog
		Private panel1 As System.Windows.Forms.Panel
		Private label1 As System.Windows.Forms.Label
		Private edFileName As System.Windows.Forms.TextBox
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
		Private WithEvents chExportAll As System.Windows.Forms.CheckBox
		Private exportDialog As System.Windows.Forms.SaveFileDialog
		Private panel4 As System.Windows.Forms.Panel
		Private label19 As System.Windows.Forms.Label
		Private chEmbed As System.Windows.Forms.CheckBox
		Private label20 As System.Windows.Forms.Label
		Private cbFontMapping As System.Windows.Forms.ComboBox
		Private panel5 As System.Windows.Forms.Panel
		Private label4 As System.Windows.Forms.Label
		Private edZoom As System.Windows.Forms.TextBox
		Private WithEvents chFitIn As System.Windows.Forms.CheckBox
		Private label6 As System.Windows.Forms.Label
		Private label5 As System.Windows.Forms.Label
		Private edVPages As System.Windows.Forms.TextBox
		Private edHPages As System.Windows.Forms.TextBox
		Private label21 As System.Windows.Forms.Label
		Private panel6 As System.Windows.Forms.Panel
		Private label11 As System.Windows.Forms.Label
		Private edf As System.Windows.Forms.TextBox
		Private labelb As System.Windows.Forms.Label
		Private edb As System.Windows.Forms.TextBox
		Private label9 As System.Windows.Forms.Label
		Private edr As System.Windows.Forms.TextBox
		Private label8 As System.Windows.Forms.Label
		Private edt As System.Windows.Forms.TextBox
		Private label7 As System.Windows.Forms.Label
		Private edl As System.Windows.Forms.TextBox
		Private label10 As System.Windows.Forms.Label
		Private edh As System.Windows.Forms.TextBox
		Private label22 As System.Windows.Forms.Label
		Private panel7 As System.Windows.Forms.Panel
		Private label23 As System.Windows.Forms.Label
		Private edFooter As System.Windows.Forms.TextBox
		Private label3 As System.Windows.Forms.Label
		Private edHeader As System.Windows.Forms.TextBox
		Private label2 As System.Windows.Forms.Label
		Private panel8 As System.Windows.Forms.Panel
		Private chPrintLeft As System.Windows.Forms.CheckBox
		Private chFormulaText As System.Windows.Forms.CheckBox
		Private chGridLines As System.Windows.Forms.CheckBox
		Private label24 As System.Windows.Forms.Label
		Private panel9 As System.Windows.Forms.Panel
		Private label25 As System.Windows.Forms.Label
		Private label26 As System.Windows.Forms.Label
		Private edAuthor As System.Windows.Forms.TextBox
		Private label27 As System.Windows.Forms.Label
		Private label28 As System.Windows.Forms.Label
		Private edTitle As System.Windows.Forms.TextBox
		Private edSubject As System.Windows.Forms.TextBox
		Private cbKerning As System.Windows.Forms.CheckBox
		Private chLandscape As System.Windows.Forms.CheckBox
		Private cbResetPageNumber As System.Windows.Forms.CheckBox
		Private cbUseGetFontData As System.Windows.Forms.CheckBox
		Private cbConfidential As System.Windows.Forms.CheckBox
		Private chSubset As System.Windows.Forms.CheckBox
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
			Dim tPdfProperties1 As New FlexCel.Pdf.TPdfProperties()
			Me.openFileDialog1 = New System.Windows.Forms.OpenFileDialog()
			Me.panel1 = New System.Windows.Forms.Panel()
			Me.panel2 = New System.Windows.Forms.Panel()
			Me.label32 = New System.Windows.Forms.Label()
			Me.label31 = New System.Windows.Forms.Label()
			Me.label30 = New System.Windows.Forms.Label()
			Me.cbTagged = New System.Windows.Forms.ComboBox()
			Me.cbVersion = New System.Windows.Forms.ComboBox()
			Me.cbPdfType = New System.Windows.Forms.ComboBox()
			Me.label34 = New System.Windows.Forms.Label()
			Me.cbConfidential = New System.Windows.Forms.CheckBox()
			Me.cbUseGetFontData = New System.Windows.Forms.CheckBox()
			Me.cbResetPageNumber = New System.Windows.Forms.CheckBox()
			Me.panel9 = New System.Windows.Forms.Panel()
			Me.label29 = New System.Windows.Forms.Label()
			Me.edLang = New System.Windows.Forms.TextBox()
			Me.edSubject = New System.Windows.Forms.TextBox()
			Me.label28 = New System.Windows.Forms.Label()
			Me.edTitle = New System.Windows.Forms.TextBox()
			Me.label27 = New System.Windows.Forms.Label()
			Me.label26 = New System.Windows.Forms.Label()
			Me.edAuthor = New System.Windows.Forms.TextBox()
			Me.label25 = New System.Windows.Forms.Label()
			Me.edFileName = New System.Windows.Forms.TextBox()
			Me.panel8 = New System.Windows.Forms.Panel()
			Me.chLandscape = New System.Windows.Forms.CheckBox()
			Me.label24 = New System.Windows.Forms.Label()
			Me.chPrintLeft = New System.Windows.Forms.CheckBox()
			Me.chFormulaText = New System.Windows.Forms.CheckBox()
			Me.chGridLines = New System.Windows.Forms.CheckBox()
			Me.panel7 = New System.Windows.Forms.Panel()
			Me.edFooter = New System.Windows.Forms.TextBox()
			Me.label3 = New System.Windows.Forms.Label()
			Me.edHeader = New System.Windows.Forms.TextBox()
			Me.label2 = New System.Windows.Forms.Label()
			Me.label23 = New System.Windows.Forms.Label()
			Me.panel6 = New System.Windows.Forms.Panel()
			Me.label22 = New System.Windows.Forms.Label()
			Me.edf = New System.Windows.Forms.TextBox()
			Me.edb = New System.Windows.Forms.TextBox()
			Me.edr = New System.Windows.Forms.TextBox()
			Me.edt = New System.Windows.Forms.TextBox()
			Me.label7 = New System.Windows.Forms.Label()
			Me.edl = New System.Windows.Forms.TextBox()
			Me.edh = New System.Windows.Forms.TextBox()
			Me.label9 = New System.Windows.Forms.Label()
			Me.label10 = New System.Windows.Forms.Label()
			Me.label8 = New System.Windows.Forms.Label()
			Me.labelb = New System.Windows.Forms.Label()
			Me.label11 = New System.Windows.Forms.Label()
			Me.panel5 = New System.Windows.Forms.Panel()
			Me.label21 = New System.Windows.Forms.Label()
			Me.label4 = New System.Windows.Forms.Label()
			Me.edZoom = New System.Windows.Forms.TextBox()
			Me.chFitIn = New System.Windows.Forms.CheckBox()
			Me.label6 = New System.Windows.Forms.Label()
			Me.label5 = New System.Windows.Forms.Label()
			Me.edVPages = New System.Windows.Forms.TextBox()
			Me.edHPages = New System.Windows.Forms.TextBox()
			Me.panel4 = New System.Windows.Forms.Panel()
			Me.chSubset = New System.Windows.Forms.CheckBox()
			Me.cbKerning = New System.Windows.Forms.CheckBox()
			Me.label20 = New System.Windows.Forms.Label()
			Me.cbFontMapping = New System.Windows.Forms.ComboBox()
			Me.chEmbed = New System.Windows.Forms.CheckBox()
			Me.label19 = New System.Windows.Forms.Label()
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
			Me.chExportAll = New System.Windows.Forms.CheckBox()
			Me.label1 = New System.Windows.Forms.Label()
			Me.exportDialog = New System.Windows.Forms.SaveFileDialog()
			Me.mainToolbar = New System.Windows.Forms.ToolStrip()
			Me.openFile = New System.Windows.Forms.ToolStripButton()
			Me.export = New System.Windows.Forms.ToolStripButton()
			Me.btnClose = New System.Windows.Forms.ToolStripButton()
			Me.flexCelPdfExport1 = New FlexCel.Render.FlexCelPdfExport()
			Me.panel1.SuspendLayout()
			Me.panel2.SuspendLayout()
			Me.panel9.SuspendLayout()
			Me.panel8.SuspendLayout()
			Me.panel7.SuspendLayout()
			Me.panel6.SuspendLayout()
			Me.panel5.SuspendLayout()
			Me.panel4.SuspendLayout()
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
			' panel1
			' 
			Me.panel1.BackColor = System.Drawing.Color.White
			Me.panel1.Controls.Add(Me.panel2)
			Me.panel1.Controls.Add(Me.cbConfidential)
			Me.panel1.Controls.Add(Me.cbUseGetFontData)
			Me.panel1.Controls.Add(Me.cbResetPageNumber)
			Me.panel1.Controls.Add(Me.panel9)
			Me.panel1.Controls.Add(Me.edFileName)
			Me.panel1.Controls.Add(Me.panel8)
			Me.panel1.Controls.Add(Me.panel7)
			Me.panel1.Controls.Add(Me.panel6)
			Me.panel1.Controls.Add(Me.panel5)
			Me.panel1.Controls.Add(Me.panel4)
			Me.panel1.Controls.Add(Me.label18)
			Me.panel1.Controls.Add(Me.cbSheet)
			Me.panel1.Controls.Add(Me.panel3)
			Me.panel1.Controls.Add(Me.chExportAll)
			Me.panel1.Controls.Add(Me.label1)
			Me.panel1.Dock = System.Windows.Forms.DockStyle.Fill
			Me.panel1.Location = New System.Drawing.Point(0, 38)
			Me.panel1.Name = "panel1"
			Me.panel1.Size = New System.Drawing.Size(768, 593)
			Me.panel1.TabIndex = 3
			' 
			' panel2
			' 
			Me.panel2.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.panel2.BackColor = System.Drawing.Color.FromArgb((CInt((CByte(192)))), (CInt((CByte(192)))), (CInt((CByte(255)))))
			Me.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.panel2.Controls.Add(Me.label32)
			Me.panel2.Controls.Add(Me.label31)
			Me.panel2.Controls.Add(Me.label30)
			Me.panel2.Controls.Add(Me.cbTagged)
			Me.panel2.Controls.Add(Me.cbVersion)
			Me.panel2.Controls.Add(Me.cbPdfType)
			Me.panel2.Controls.Add(Me.label34)
			Me.panel2.Location = New System.Drawing.Point(32, 146)
			Me.panel2.Name = "panel2"
			Me.panel2.Size = New System.Drawing.Size(688, 53)
			Me.panel2.TabIndex = 42
			' 
			' label32
			' 
			Me.label32.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label32.Location = New System.Drawing.Point(8, 22)
			Me.label32.Name = "label32"
			Me.label32.Size = New System.Drawing.Size(41, 16)
			Me.label32.TabIndex = 41
			Me.label32.Text = "Type:"
			' 
			' label31
			' 
			Me.label31.AutoSize = True
			Me.label31.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label31.Location = New System.Drawing.Point(219, 22)
			Me.label31.Name = "label31"
			Me.label31.Size = New System.Drawing.Size(53, 14)
			Me.label31.TabIndex = 40
			Me.label31.Text = "Version:"
			' 
			' label30
			' 
			Me.label30.AutoSize = True
			Me.label30.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label30.Location = New System.Drawing.Point(467, 22)
			Me.label30.Name = "label30"
			Me.label30.Size = New System.Drawing.Size(50, 14)
			Me.label30.TabIndex = 39
			Me.label30.Text = "Tagged:"
			' 
			' cbTagged
			' 
			Me.cbTagged.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
			Me.cbTagged.Items.AddRange(New Object() { "Full", "None"})
			Me.cbTagged.Location = New System.Drawing.Point(523, 19)
			Me.cbTagged.Name = "cbTagged"
			Me.cbTagged.Size = New System.Drawing.Size(149, 21)
			Me.cbTagged.TabIndex = 36
			' 
			' cbVersion
			' 
			Me.cbVersion.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
			Me.cbVersion.Items.AddRange(New Object() { "1.4 (Acrobat 5)", "1.6 (Acrobat 7)"})
			Me.cbVersion.Location = New System.Drawing.Point(278, 19)
			Me.cbVersion.Name = "cbVersion"
			Me.cbVersion.Size = New System.Drawing.Size(168, 21)
			Me.cbVersion.TabIndex = 35
			' 
			' cbPdfType
			' 
			Me.cbPdfType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
			Me.cbPdfType.Items.AddRange(New Object() { "Standard", "PDF/A1", "PDF/A2", "PDF/A3"})
			Me.cbPdfType.Location = New System.Drawing.Point(56, 19)
			Me.cbPdfType.Name = "cbPdfType"
			Me.cbPdfType.Size = New System.Drawing.Size(144, 21)
			Me.cbPdfType.TabIndex = 34
			' 
			' label34
			' 
			Me.label34.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label34.Location = New System.Drawing.Point(8, 0)
			Me.label34.Name = "label34"
			Me.label34.Size = New System.Drawing.Size(192, 16)
			Me.label34.TabIndex = 20
			Me.label34.Text = "Pdf options:"
			' 
			' cbConfidential
			' 
			Me.cbConfidential.Anchor = (CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.cbConfidential.Location = New System.Drawing.Point(488, 557)
			Me.cbConfidential.Name = "cbConfidential"
			Me.cbConfidential.Size = New System.Drawing.Size(232, 16)
			Me.cbConfidential.TabIndex = 41
			Me.cbConfidential.Text = "Print ""Confidential"" on each page"
			' 
			' cbUseGetFontData
			' 
			Me.cbUseGetFontData.Location = New System.Drawing.Point(32, 557)
			Me.cbUseGetFontData.Name = "cbUseGetFontData"
			Me.cbUseGetFontData.Size = New System.Drawing.Size(312, 16)
			Me.cbUseGetFontData.TabIndex = 40
			Me.cbUseGetFontData.Text = "Use UNMANAGED calls to Win32 API to find the fonts."
			' 
			' cbResetPageNumber
			' 
			Me.cbResetPageNumber.Anchor = (CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.cbResetPageNumber.Location = New System.Drawing.Point(528, 40)
			Me.cbResetPageNumber.Name = "cbResetPageNumber"
			Me.cbResetPageNumber.Size = New System.Drawing.Size(200, 16)
			Me.cbResetPageNumber.TabIndex = 39
			Me.cbResetPageNumber.Text = "Reset Page number on each sheet"
			' 
			' panel9
			' 
			Me.panel9.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.panel9.BackColor = System.Drawing.Color.FromArgb((CInt((CByte(255)))), (CInt((CByte(192)))), (CInt((CByte(255)))))
			Me.panel9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.panel9.Controls.Add(Me.label29)
			Me.panel9.Controls.Add(Me.edLang)
			Me.panel9.Controls.Add(Me.edSubject)
			Me.panel9.Controls.Add(Me.label28)
			Me.panel9.Controls.Add(Me.edTitle)
			Me.panel9.Controls.Add(Me.label27)
			Me.panel9.Controls.Add(Me.label26)
			Me.panel9.Controls.Add(Me.edAuthor)
			Me.panel9.Controls.Add(Me.label25)
			Me.panel9.Location = New System.Drawing.Point(32, 64)
			Me.panel9.Name = "panel9"
			Me.panel9.Size = New System.Drawing.Size(688, 76)
			Me.panel9.TabIndex = 38
			' 
			' label29
			' 
			Me.label29.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label29.Location = New System.Drawing.Point(8, 49)
			Me.label29.Name = "label29"
			Me.label29.Size = New System.Drawing.Size(48, 16)
			Me.label29.TabIndex = 38
			Me.label29.Text = "Lang:"
			' 
			' edLang
			' 
			Me.edLang.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.edLang.Location = New System.Drawing.Point(56, 46)
			Me.edLang.Name = "edLang"
			Me.edLang.Size = New System.Drawing.Size(144, 20)
			Me.edLang.TabIndex = 37
			Me.edLang.Text = "en-US"
			' 
			' edSubject
			' 
			Me.edSubject.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.edSubject.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.edSubject.Location = New System.Drawing.Point(278, 45)
			Me.edSubject.Name = "edSubject"
			Me.edSubject.Size = New System.Drawing.Size(394, 20)
			Me.edSubject.TabIndex = 35
			' 
			' label28
			' 
			Me.label28.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label28.Location = New System.Drawing.Point(216, 50)
			Me.label28.Name = "label28"
			Me.label28.Size = New System.Drawing.Size(56, 16)
			Me.label28.TabIndex = 36
			Me.label28.Text = "Subject:"
			' 
			' edTitle
			' 
			Me.edTitle.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.edTitle.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.edTitle.Location = New System.Drawing.Point(278, 19)
			Me.edTitle.Name = "edTitle"
			Me.edTitle.Size = New System.Drawing.Size(394, 20)
			Me.edTitle.TabIndex = 33
			' 
			' label27
			' 
			Me.label27.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label27.Location = New System.Drawing.Point(232, 22)
			Me.label27.Name = "label27"
			Me.label27.Size = New System.Drawing.Size(48, 16)
			Me.label27.TabIndex = 34
			Me.label27.Text = "Title:"
			' 
			' label26
			' 
			Me.label26.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label26.Location = New System.Drawing.Point(8, 23)
			Me.label26.Name = "label26"
			Me.label26.Size = New System.Drawing.Size(48, 16)
			Me.label26.TabIndex = 32
			Me.label26.Text = "Author:"
			' 
			' edAuthor
			' 
			Me.edAuthor.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.edAuthor.Location = New System.Drawing.Point(56, 20)
			Me.edAuthor.Name = "edAuthor"
			Me.edAuthor.Size = New System.Drawing.Size(144, 20)
			Me.edAuthor.TabIndex = 31
			' 
			' label25
			' 
			Me.label25.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label25.Location = New System.Drawing.Point(8, 0)
			Me.label25.Name = "label25"
			Me.label25.Size = New System.Drawing.Size(192, 16)
			Me.label25.TabIndex = 20
			Me.label25.Text = "Pdf Properties:"
			' 
			' edFileName
			' 
			Me.edFileName.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.edFileName.BackColor = System.Drawing.Color.White
			Me.edFileName.BorderStyle = System.Windows.Forms.BorderStyle.None
			Me.edFileName.Location = New System.Drawing.Point(136, 16)
			Me.edFileName.Name = "edFileName"
			Me.edFileName.ReadOnly = True
			Me.edFileName.Size = New System.Drawing.Size(584, 13)
			Me.edFileName.TabIndex = 4
			Me.edFileName.Text = "No file selected"
			' 
			' panel8
			' 
			Me.panel8.BackColor = System.Drawing.Color.FromArgb((CInt((CByte(224)))), (CInt((CByte(224)))), (CInt((CByte(224)))))
			Me.panel8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.panel8.Controls.Add(Me.chLandscape)
			Me.panel8.Controls.Add(Me.label24)
			Me.panel8.Controls.Add(Me.chPrintLeft)
			Me.panel8.Controls.Add(Me.chFormulaText)
			Me.panel8.Controls.Add(Me.chGridLines)
			Me.panel8.Location = New System.Drawing.Point(32, 205)
			Me.panel8.Name = "panel8"
			Me.panel8.Size = New System.Drawing.Size(176, 120)
			Me.panel8.TabIndex = 37
			' 
			' chLandscape
			' 
			Me.chLandscape.Location = New System.Drawing.Point(24, 88)
			Me.chLandscape.Name = "chLandscape"
			Me.chLandscape.Size = New System.Drawing.Size(136, 24)
			Me.chLandscape.TabIndex = 20
			Me.chLandscape.Text = "Landscape"
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
			' chPrintLeft
			' 
			Me.chPrintLeft.Location = New System.Drawing.Point(24, 47)
			Me.chPrintLeft.Name = "chPrintLeft"
			Me.chPrintLeft.Size = New System.Drawing.Size(152, 16)
			Me.chPrintLeft.TabIndex = 18
			Me.chPrintLeft.Text = "Print Left, then down."
			' 
			' chFormulaText
			' 
			Me.chFormulaText.Location = New System.Drawing.Point(24, 71)
			Me.chFormulaText.Name = "chFormulaText"
			Me.chFormulaText.Size = New System.Drawing.Size(136, 16)
			Me.chFormulaText.TabIndex = 17
			Me.chFormulaText.Text = "Print Formula Text"
			' 
			' chGridLines
			' 
			Me.chGridLines.Location = New System.Drawing.Point(24, 24)
			Me.chGridLines.Name = "chGridLines"
			Me.chGridLines.Size = New System.Drawing.Size(128, 16)
			Me.chGridLines.TabIndex = 16
			Me.chGridLines.Text = "Print Grid Lines"
			' 
			' panel7
			' 
			Me.panel7.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.panel7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.panel7.Controls.Add(Me.edFooter)
			Me.panel7.Controls.Add(Me.label3)
			Me.panel7.Controls.Add(Me.edHeader)
			Me.panel7.Controls.Add(Me.label2)
			Me.panel7.Controls.Add(Me.label23)
			Me.panel7.Location = New System.Drawing.Point(224, 333)
			Me.panel7.Name = "panel7"
			Me.panel7.Size = New System.Drawing.Size(296, 112)
			Me.panel7.TabIndex = 36
			' 
			' edFooter
			' 
			Me.edFooter.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.edFooter.BackColor = System.Drawing.Color.White
			Me.edFooter.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.edFooter.Location = New System.Drawing.Point(8, 88)
			Me.edFooter.Name = "edFooter"
			Me.edFooter.Size = New System.Drawing.Size(278, 20)
			Me.edFooter.TabIndex = 46
			' 
			' label3
			' 
			Me.label3.BackColor = System.Drawing.Color.White
			Me.label3.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label3.Location = New System.Drawing.Point(8, 72)
			Me.label3.Name = "label3"
			Me.label3.Size = New System.Drawing.Size(56, 16)
			Me.label3.TabIndex = 45
			Me.label3.Text = "Footer:"
			' 
			' edHeader
			' 
			Me.edHeader.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.edHeader.BackColor = System.Drawing.Color.White
			Me.edHeader.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.edHeader.Location = New System.Drawing.Point(8, 48)
			Me.edHeader.Name = "edHeader"
			Me.edHeader.Size = New System.Drawing.Size(278, 20)
			Me.edHeader.TabIndex = 44
			' 
			' label2
			' 
			Me.label2.BackColor = System.Drawing.Color.White
			Me.label2.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label2.Location = New System.Drawing.Point(8, 32)
			Me.label2.Name = "label2"
			Me.label2.Size = New System.Drawing.Size(56, 16)
			Me.label2.TabIndex = 43
			Me.label2.Text = "Header:"
			' 
			' label23
			' 
			Me.label23.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label23.Location = New System.Drawing.Point(8, 8)
			Me.label23.Name = "label23"
			Me.label23.Size = New System.Drawing.Size(187, 16)
			Me.label23.TabIndex = 42
			Me.label23.Text = "Headers and footers:"
			' 
			' panel6
			' 
			Me.panel6.BackColor = System.Drawing.Color.FromArgb((CInt((CByte(255)))), (CInt((CByte(224)))), (CInt((CByte(192)))))
			Me.panel6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.panel6.Controls.Add(Me.label22)
			Me.panel6.Controls.Add(Me.edf)
			Me.panel6.Controls.Add(Me.edb)
			Me.panel6.Controls.Add(Me.edr)
			Me.panel6.Controls.Add(Me.edt)
			Me.panel6.Controls.Add(Me.label7)
			Me.panel6.Controls.Add(Me.edl)
			Me.panel6.Controls.Add(Me.edh)
			Me.panel6.Controls.Add(Me.label9)
			Me.panel6.Controls.Add(Me.label10)
			Me.panel6.Controls.Add(Me.label8)
			Me.panel6.Controls.Add(Me.labelb)
			Me.panel6.Controls.Add(Me.label11)
			Me.panel6.Location = New System.Drawing.Point(32, 333)
			Me.panel6.Name = "panel6"
			Me.panel6.Size = New System.Drawing.Size(176, 208)
			Me.panel6.TabIndex = 35
			' 
			' label22
			' 
			Me.label22.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label22.Location = New System.Drawing.Point(8, 8)
			Me.label22.Name = "label22"
			Me.label22.Size = New System.Drawing.Size(120, 16)
			Me.label22.TabIndex = 41
			Me.label22.Text = "Margins:"
			' 
			' edf
			' 
			Me.edf.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.edf.Location = New System.Drawing.Point(56, 152)
			Me.edf.Name = "edf"
			Me.edf.Size = New System.Drawing.Size(112, 20)
			Me.edf.TabIndex = 39
			' 
			' edb
			' 
			Me.edb.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.edb.Location = New System.Drawing.Point(56, 128)
			Me.edb.Name = "edb"
			Me.edb.Size = New System.Drawing.Size(112, 20)
			Me.edb.TabIndex = 37
			' 
			' edr
			' 
			Me.edr.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.edr.Location = New System.Drawing.Point(56, 56)
			Me.edr.Name = "edr"
			Me.edr.Size = New System.Drawing.Size(112, 20)
			Me.edr.TabIndex = 35
			' 
			' edt
			' 
			Me.edt.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.edt.Location = New System.Drawing.Point(56, 104)
			Me.edt.Name = "edt"
			Me.edt.Size = New System.Drawing.Size(112, 20)
			Me.edt.TabIndex = 31
			' 
			' label7
			' 
			Me.label7.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label7.Location = New System.Drawing.Point(8, 32)
			Me.label7.Name = "label7"
			Me.label7.Size = New System.Drawing.Size(36, 16)
			Me.label7.TabIndex = 30
			Me.label7.Text = "Left:"
			' 
			' edl
			' 
			Me.edl.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.edl.Location = New System.Drawing.Point(56, 32)
			Me.edl.Name = "edl"
			Me.edl.Size = New System.Drawing.Size(112, 20)
			Me.edl.TabIndex = 29
			' 
			' edh
			' 
			Me.edh.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.edh.Location = New System.Drawing.Point(56, 80)
			Me.edh.Name = "edh"
			Me.edh.Size = New System.Drawing.Size(112, 20)
			Me.edh.TabIndex = 32
			' 
			' label9
			' 
			Me.label9.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label9.Location = New System.Drawing.Point(8, 56)
			Me.label9.Name = "label9"
			Me.label9.Size = New System.Drawing.Size(80, 16)
			Me.label9.TabIndex = 36
			Me.label9.Text = "Right:"
			' 
			' label10
			' 
			Me.label10.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label10.Location = New System.Drawing.Point(7, 80)
			Me.label10.Name = "label10"
			Me.label10.Size = New System.Drawing.Size(88, 16)
			Me.label10.TabIndex = 34
			Me.label10.Text = "Header:"
			' 
			' label8
			' 
			Me.label8.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label8.Location = New System.Drawing.Point(8, 104)
			Me.label8.Name = "label8"
			Me.label8.Size = New System.Drawing.Size(80, 16)
			Me.label8.TabIndex = 33
			Me.label8.Text = "Top:"
			' 
			' labelb
			' 
			Me.labelb.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.labelb.Location = New System.Drawing.Point(8, 130)
			Me.labelb.Name = "labelb"
			Me.labelb.Size = New System.Drawing.Size(88, 16)
			Me.labelb.TabIndex = 38
			Me.labelb.Text = "Bottom:"
			' 
			' label11
			' 
			Me.label11.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label11.Location = New System.Drawing.Point(8, 160)
			Me.label11.Name = "label11"
			Me.label11.Size = New System.Drawing.Size(80, 16)
			Me.label11.TabIndex = 40
			Me.label11.Text = "Footer:"
			' 
			' panel5
			' 
			Me.panel5.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.panel5.BackColor = System.Drawing.Color.FromArgb((CInt((CByte(192)))), (CInt((CByte(255)))), (CInt((CByte(192)))))
			Me.panel5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.panel5.Controls.Add(Me.label21)
			Me.panel5.Controls.Add(Me.label4)
			Me.panel5.Controls.Add(Me.edZoom)
			Me.panel5.Controls.Add(Me.chFitIn)
			Me.panel5.Controls.Add(Me.label6)
			Me.panel5.Controls.Add(Me.label5)
			Me.panel5.Controls.Add(Me.edVPages)
			Me.panel5.Controls.Add(Me.edHPages)
			Me.panel5.Location = New System.Drawing.Point(224, 453)
			Me.panel5.Name = "panel5"
			Me.panel5.Size = New System.Drawing.Size(296, 88)
			Me.panel5.TabIndex = 34
			' 
			' label21
			' 
			Me.label21.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label21.Location = New System.Drawing.Point(9, 5)
			Me.label21.Name = "label21"
			Me.label21.Size = New System.Drawing.Size(192, 16)
			Me.label21.TabIndex = 26
			Me.label21.Text = "Zoom:"
			' 
			' label4
			' 
			Me.label4.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label4.Location = New System.Drawing.Point(72, 58)
			Me.label4.Name = "label4"
			Me.label4.Size = New System.Drawing.Size(56, 16)
			Me.label4.TabIndex = 25
			Me.label4.Text = "Zoom (%)"
			' 
			' edZoom
			' 
			Me.edZoom.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.edZoom.Location = New System.Drawing.Point(136, 56)
			Me.edZoom.Name = "edZoom"
			Me.edZoom.Size = New System.Drawing.Size(24, 20)
			Me.edZoom.TabIndex = 24
			' 
			' chFitIn
			' 
			Me.chFitIn.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.chFitIn.Location = New System.Drawing.Point(8, 24)
			Me.chFitIn.Name = "chFitIn"
			Me.chFitIn.Size = New System.Drawing.Size(56, 24)
			Me.chFitIn.TabIndex = 23
			Me.chFitIn.Text = "Fit in"
'			Me.chFitIn.CheckedChanged += New System.EventHandler(Me.chFitIn_CheckedChanged)
			' 
			' label6
			' 
			Me.label6.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label6.Location = New System.Drawing.Point(208, 29)
			Me.label6.Name = "label6"
			Me.label6.Size = New System.Drawing.Size(80, 16)
			Me.label6.TabIndex = 22
			Me.label6.Text = "pages tall."
			' 
			' label5
			' 
			Me.label5.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label5.Location = New System.Drawing.Point(96, 29)
			Me.label5.Name = "label5"
			Me.label5.Size = New System.Drawing.Size(80, 16)
			Me.label5.TabIndex = 21
			Me.label5.Text = "pages wide x"
			' 
			' edVPages
			' 
			Me.edVPages.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.edVPages.Location = New System.Drawing.Point(176, 24)
			Me.edVPages.Name = "edVPages"
			Me.edVPages.ReadOnly = True
			Me.edVPages.Size = New System.Drawing.Size(24, 20)
			Me.edVPages.TabIndex = 20
			' 
			' edHPages
			' 
			Me.edHPages.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.edHPages.Location = New System.Drawing.Point(64, 24)
			Me.edHPages.Name = "edHPages"
			Me.edHPages.ReadOnly = True
			Me.edHPages.Size = New System.Drawing.Size(24, 20)
			Me.edHPages.TabIndex = 19
			' 
			' panel4
			' 
			Me.panel4.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.panel4.BackColor = System.Drawing.Color.FromArgb((CInt((CByte(192)))), (CInt((CByte(255)))), (CInt((CByte(255)))))
			Me.panel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.panel4.Controls.Add(Me.chSubset)
			Me.panel4.Controls.Add(Me.cbKerning)
			Me.panel4.Controls.Add(Me.label20)
			Me.panel4.Controls.Add(Me.cbFontMapping)
			Me.panel4.Controls.Add(Me.chEmbed)
			Me.panel4.Controls.Add(Me.label19)
			Me.panel4.Location = New System.Drawing.Point(224, 205)
			Me.panel4.Name = "panel4"
			Me.panel4.Size = New System.Drawing.Size(496, 120)
			Me.panel4.TabIndex = 33
			' 
			' chSubset
			' 
			Me.chSubset.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.chSubset.Checked = True
			Me.chSubset.CheckState = System.Windows.Forms.CheckState.Checked
			Me.chSubset.Location = New System.Drawing.Point(16, 72)
			Me.chSubset.Name = "chSubset"
			Me.chSubset.Size = New System.Drawing.Size(464, 16)
			Me.chSubset.TabIndex = 36
			Me.chSubset.Text = "Subset fonts when embedding. (That is, embed only the characters used from the fo" & "nt)"
			' 
			' cbKerning
			' 
			Me.cbKerning.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.cbKerning.Location = New System.Drawing.Point(16, 88)
			Me.cbKerning.Name = "cbKerning"
			Me.cbKerning.Size = New System.Drawing.Size(464, 32)
			Me.cbKerning.TabIndex = 35
			Me.cbKerning.Text = "Kerning. (Files with kerning look a little better but are a little bigger too)"
			' 
			' label20
			' 
			Me.label20.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label20.Location = New System.Drawing.Point(16, 24)
			Me.label20.Name = "label20"
			Me.label20.Size = New System.Drawing.Size(96, 16)
			Me.label20.TabIndex = 34
			Me.label20.Text = "Font mapping:"
			' 
			' cbFontMapping
			' 
			Me.cbFontMapping.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.cbFontMapping.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
			Me.cbFontMapping.Items.AddRange(New Object() { "Replace all fonts by internal fonts. (smaller file size)", "Replace selected fonts by internal fonts. (optimum relation file size/accuracy)", "Do not replace any font. (maximum file size)"})
			Me.cbFontMapping.Location = New System.Drawing.Point(120, 24)
			Me.cbFontMapping.Name = "cbFontMapping"
			Me.cbFontMapping.Size = New System.Drawing.Size(360, 21)
			Me.cbFontMapping.TabIndex = 33
			' 
			' chEmbed
			' 
			Me.chEmbed.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.chEmbed.Checked = True
			Me.chEmbed.CheckState = System.Windows.Forms.CheckState.Checked
			Me.chEmbed.Location = New System.Drawing.Point(16, 40)
			Me.chEmbed.Name = "chEmbed"
			Me.chEmbed.Size = New System.Drawing.Size(464, 32)
			Me.chEmbed.TabIndex = 3
			Me.chEmbed.Text = "Embed all fonts. (if you leave this option off, some fonts might be embedded anyw" & "ay)"
			' 
			' label19
			' 
			Me.label19.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label19.Location = New System.Drawing.Point(3, 8)
			Me.label19.Name = "label19"
			Me.label19.Size = New System.Drawing.Size(192, 16)
			Me.label19.TabIndex = 2
			Me.label19.Text = "Fonts:"
			' 
			' label18
			' 
			Me.label18.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label18.Location = New System.Drawing.Point(168, 40)
			Me.label18.Name = "label18"
			Me.label18.Size = New System.Drawing.Size(96, 16)
			Me.label18.TabIndex = 32
			Me.label18.Text = "Sheet to export:"
			' 
			' cbSheet
			' 
			Me.cbSheet.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.cbSheet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
			Me.cbSheet.Location = New System.Drawing.Point(264, 34)
			Me.cbSheet.Name = "cbSheet"
			Me.cbSheet.Size = New System.Drawing.Size(256, 21)
			Me.cbSheet.TabIndex = 31
'			Me.cbSheet.SelectedIndexChanged += New System.EventHandler(Me.cbSheet_SelectedIndexChanged)
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
			Me.panel3.Location = New System.Drawing.Point(536, 333)
			Me.panel3.Name = "panel3"
			Me.panel3.Size = New System.Drawing.Size(184, 208)
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
			Me.label13.Size = New System.Drawing.Size(160, 32)
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
			Me.label12.Text = "Range to Export:"
			' 
			' chExportAll
			' 
			Me.chExportAll.Location = New System.Drawing.Point(32, 40)
			Me.chExportAll.Name = "chExportAll"
			Me.chExportAll.Size = New System.Drawing.Size(128, 16)
			Me.chExportAll.TabIndex = 1
			Me.chExportAll.Text = "Export all Sheets"
'			Me.chExportAll.CheckedChanged += New System.EventHandler(Me.chExportAll_CheckedChanged)
			' 
			' label1
			' 
			Me.label1.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label1.Location = New System.Drawing.Point(29, 16)
			Me.label1.Name = "label1"
			Me.label1.Size = New System.Drawing.Size(80, 16)
			Me.label1.TabIndex = 0
			Me.label1.Text = "File to print:"
			' 
			' exportDialog
			' 
			Me.exportDialog.DefaultExt = "pdf"
			Me.exportDialog.Filter = "Pdf files|*.pdf"
			' 
			' mainToolbar
			' 
			Me.mainToolbar.Items.AddRange(New System.Windows.Forms.ToolStripItem() { Me.openFile, Me.export, Me.btnClose})
			Me.mainToolbar.Location = New System.Drawing.Point(0, 0)
			Me.mainToolbar.Name = "mainToolbar"
			Me.mainToolbar.Size = New System.Drawing.Size(768, 38)
			Me.mainToolbar.TabIndex = 4
			Me.mainToolbar.Text = "toolStrip1"
			' 
			' openFile
			' 
			Me.openFile.Image = (CType(resources.GetObject("openFile.Image"), System.Drawing.Image))
			Me.openFile.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.openFile.Name = "openFile"
			Me.openFile.Size = New System.Drawing.Size(61, 35)
			Me.openFile.Text = "Open File"
			Me.openFile.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
'			Me.openFile.Click += New System.EventHandler(Me.openFile_Click)
			' 
			' export
			' 
			Me.export.Image = (CType(resources.GetObject("export.Image"), System.Drawing.Image))
			Me.export.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.export.Name = "export"
			Me.export.Size = New System.Drawing.Size(82, 35)
			Me.export.Text = "Export to PDF"
			Me.export.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
'			Me.export.Click += New System.EventHandler(Me.export_Click)
			' 
			' btnClose
			' 
			Me.btnClose.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
			Me.btnClose.Image = (CType(resources.GetObject("btnClose.Image"), System.Drawing.Image))
			Me.btnClose.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnClose.Name = "btnClose"
			Me.btnClose.Size = New System.Drawing.Size(59, 35)
			Me.btnClose.Text = "     E&xit     "
			Me.btnClose.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
'			Me.btnClose.Click += New System.EventHandler(Me.button2_Click)
			' 
			' flexCelPdfExport1
			' 
			Me.flexCelPdfExport1.FontEmbed = FlexCel.Pdf.TFontEmbed.Embed
			Me.flexCelPdfExport1.InitialZoomAndPage = Nothing
			Me.flexCelPdfExport1.PageLayout = FlexCel.Pdf.TPageLayout.None
			Me.flexCelPdfExport1.PageLayoutDisplay = FlexCel.Pdf.TPageLayoutDisplay.None
			Me.flexCelPdfExport1.PageSize = Nothing
			tPdfProperties1.Author = Nothing
			tPdfProperties1.Creator = Nothing
			tPdfProperties1.Keywords = Nothing
			tPdfProperties1.Language = Nothing
			tPdfProperties1.Subject = Nothing
			tPdfProperties1.Title = Nothing
			Me.flexCelPdfExport1.Properties = tPdfProperties1
			Me.flexCelPdfExport1.TagMode = FlexCel.Pdf.TTagMode.Full
			Me.flexCelPdfExport1.UnlicensedReplacementFont = Nothing
			Me.flexCelPdfExport1.UseExcelProperties = True
			Me.flexCelPdfExport1.Workbook = Nothing
'			Me.flexCelPdfExport1.AfterGeneratePage += New FlexCel.Render.PageEventHandler(Me.flexCelPdfExport1_AfterGeneratePage)
'			Me.flexCelPdfExport1.GetFontData += New FlexCel.Pdf.GetFontDataEventHandler(Me.flexCelPdfExport1_GetFontData)
			' 
			' mainForm
			' 
			Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New System.Drawing.Size(768, 631)
			Me.Controls.Add(Me.panel1)
			Me.Controls.Add(Me.mainToolbar)
			Me.Name = "mainForm"
			Me.Text = "Export an Excel file to pdf"
			Me.panel1.ResumeLayout(False)
			Me.panel1.PerformLayout()
			Me.panel2.ResumeLayout(False)
			Me.panel2.PerformLayout()
			Me.panel9.ResumeLayout(False)
			Me.panel9.PerformLayout()
			Me.panel8.ResumeLayout(False)
			Me.panel7.ResumeLayout(False)
			Me.panel7.PerformLayout()
			Me.panel6.ResumeLayout(False)
			Me.panel6.PerformLayout()
			Me.panel5.ResumeLayout(False)
			Me.panel5.PerformLayout()
			Me.panel4.ResumeLayout(False)
			Me.panel3.ResumeLayout(False)
			Me.panel3.PerformLayout()
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
		Private label29 As Label
		Private edLang As TextBox
		Private panel2 As Panel
		Private cbPdfType As ComboBox
		Private label34 As Label
		Private cbVersion As ComboBox
		Private label31 As Label
		Private label30 As Label
		Private cbTagged As ComboBox
		Private label32 As Label
	End Class
End Namespace

