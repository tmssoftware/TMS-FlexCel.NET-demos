Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports System.Xml
Imports System.Net
Imports System.Threading
Imports System.Globalization
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Report
Imports FlexCel.Render
Namespace HTML
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		''' <summary>
		''' Required designer variable.
		''' </summary>
		Private components As System.ComponentModel.Container = Nothing

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
			Me.btnExportXls = New System.Windows.Forms.Button()
			Me.saveFileDialogXls = New System.Windows.Forms.SaveFileDialog()
			Me.btnCancel = New System.Windows.Forms.Button()
			Me.cbOffline = New System.Windows.Forms.CheckBox()
			Me.btnExportPdf = New System.Windows.Forms.Button()
			Me.saveFileDialogPdf = New System.Windows.Forms.SaveFileDialog()
			Me.edCity = New System.Windows.Forms.TextBox()
			Me.labelCity = New System.Windows.Forms.Label()
			Me.label1 = New System.Windows.Forms.Label()
			Me.linkLabel1 = New System.Windows.Forms.LinkLabel()
			Me.panel1 = New System.Windows.Forms.Panel()
			Me.label2 = New System.Windows.Forms.Label()
			Me.panel1.SuspendLayout()
			Me.SuspendLayout()
			' 
			' btnExportXls
			' 
			Me.btnExportXls.Anchor = (CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.btnExportXls.BackColor = System.Drawing.Color.Green
			Me.btnExportXls.ForeColor = System.Drawing.Color.White
			Me.btnExportXls.Location = New System.Drawing.Point(16, 258)
			Me.btnExportXls.Name = "btnExportXls"
			Me.btnExportXls.Size = New System.Drawing.Size(112, 24)
			Me.btnExportXls.TabIndex = 0
			Me.btnExportXls.Text = "Export to Excel"
			Me.btnExportXls.UseVisualStyleBackColor = False
'			Me.btnExportXls.Click += New System.EventHandler(Me.btnExportXls_Click)
			' 
			' saveFileDialogXls
			' 
			Me.saveFileDialogXls.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|Excel 97/2003|*.xls|Excel 2007|*.xlsx;*.xlsm|All " & "files|*.*"
			Me.saveFileDialogXls.RestoreDirectory = True
			' 
			' btnCancel
			' 
			Me.btnCancel.Anchor = (CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.btnCancel.BackColor = System.Drawing.Color.FromArgb((CInt((CByte(192)))), (CInt((CByte(0)))), (CInt((CByte(0)))))
			Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
			Me.btnCancel.ForeColor = System.Drawing.Color.White
			Me.btnCancel.Location = New System.Drawing.Point(272, 258)
			Me.btnCancel.Name = "btnCancel"
			Me.btnCancel.Size = New System.Drawing.Size(112, 24)
			Me.btnCancel.TabIndex = 3
			Me.btnCancel.Text = "Cancel"
			Me.btnCancel.UseVisualStyleBackColor = False
'			Me.btnCancel.Click += New System.EventHandler(Me.btnCancel_Click)
			' 
			' cbOffline
			' 
			Me.cbOffline.Checked = True
			Me.cbOffline.CheckState = System.Windows.Forms.CheckState.Checked
			Me.cbOffline.Location = New System.Drawing.Point(13, 200)
			Me.cbOffline.Name = "cbOffline"
			Me.cbOffline.Size = New System.Drawing.Size(352, 24)
			Me.cbOffline.TabIndex = 10
			Me.cbOffline.Text = "Use offline data. (do not actually connect to the web service)"
			' 
			' btnExportPdf
			' 
			Me.btnExportPdf.Anchor = (CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.btnExportPdf.BackColor = System.Drawing.Color.SteelBlue
			Me.btnExportPdf.ForeColor = System.Drawing.Color.White
			Me.btnExportPdf.Location = New System.Drawing.Point(144, 258)
			Me.btnExportPdf.Name = "btnExportPdf"
			Me.btnExportPdf.Size = New System.Drawing.Size(112, 24)
			Me.btnExportPdf.TabIndex = 11
			Me.btnExportPdf.Text = "Export to Pdf"
			Me.btnExportPdf.UseVisualStyleBackColor = False
'			Me.btnExportPdf.Click += New System.EventHandler(Me.btnExportPdf_Click)
			' 
			' saveFileDialogPdf
			' 
			Me.saveFileDialogPdf.Filter = "Pdf Files|*.pdf"
			Me.saveFileDialogPdf.RestoreDirectory = True
			' 
			' edCity
			' 
			Me.edCity.Location = New System.Drawing.Point(157, 160)
			Me.edCity.Name = "edCity"
			Me.edCity.Size = New System.Drawing.Size(208, 20)
			Me.edCity.TabIndex = 12
			Me.edCity.Text = "london"
			' 
			' labelCity
			' 
			Me.labelCity.Location = New System.Drawing.Point(5, 152)
			Me.labelCity.Name = "labelCity"
			Me.labelCity.Size = New System.Drawing.Size(144, 40)
			Me.labelCity.TabIndex = 13
			Me.labelCity.Text = "City Name: (try things like tokio, sydney, new york, madrid, rio de janeiro)"
			' 
			' label1
			' 
			Me.label1.Font = New System.Drawing.Font("Times New Roman", 10.25F, System.Drawing.FontStyle.Italic)
			Me.label1.Location = New System.Drawing.Point(13, 88)
			Me.label1.Name = "label1"
			Me.label1.Size = New System.Drawing.Size(352, 32)
			Me.label1.TabIndex = 14
			Me.label1.Text = "This application uses Yahoo Travel APIs to load demo trips and export them to Exc" & "el. For more information, visit:"
			' 
			' linkLabel1
			' 
			Me.linkLabel1.Location = New System.Drawing.Point(13, 128)
			Me.linkLabel1.Name = "linkLabel1"
			Me.linkLabel1.Size = New System.Drawing.Size(320, 16)
			Me.linkLabel1.TabIndex = 15
			Me.linkLabel1.TabStop = True
			Me.linkLabel1.Text = "http://developer.yahoo.com/travel/"
'			Me.linkLabel1.LinkClicked += New System.Windows.Forms.LinkLabelLinkClickedEventHandler(Me.linkLabel1_LinkClicked)
			' 
			' panel1
			' 
			Me.panel1.BackColor = System.Drawing.Color.LightGoldenrodYellow
			Me.panel1.Controls.Add(Me.label2)
			Me.panel1.Dock = System.Windows.Forms.DockStyle.Top
			Me.panel1.Location = New System.Drawing.Point(0, 0)
			Me.panel1.Name = "panel1"
			Me.panel1.Size = New System.Drawing.Size(416, 69)
			Me.panel1.TabIndex = 16
			' 
			' label2
			' 
			Me.label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label2.Location = New System.Drawing.Point(12, 9)
			Me.label2.Name = "label2"
			Me.label2.Size = New System.Drawing.Size(392, 49)
			Me.label2.TabIndex = 0
			Me.label2.Text = "IMPORTANT: Yahoo has discontinued this service, so this demo will only work with " & "offline data. As the online functionality isn't essential to this demo, we have " & "decided to keep it."
			' 
			' mainForm
			' 
			Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New System.Drawing.Size(416, 292)
			Me.Controls.Add(Me.panel1)
			Me.Controls.Add(Me.linkLabel1)
			Me.Controls.Add(Me.label1)
			Me.Controls.Add(Me.labelCity)
			Me.Controls.Add(Me.edCity)
			Me.Controls.Add(Me.btnExportPdf)
			Me.Controls.Add(Me.cbOffline)
			Me.Controls.Add(Me.btnCancel)
			Me.Controls.Add(Me.btnExportXls)
			Me.Name = "mainForm"
			Me.Text = "Using HTML formatted text with FlexCel"
			Me.panel1.ResumeLayout(False)
			Me.ResumeLayout(False)
			Me.PerformLayout()

		End Sub
		#End Region
		Private WithEvents btnCancel As System.Windows.Forms.Button
		Private cbOffline As System.Windows.Forms.CheckBox
		Private WithEvents btnExportXls As System.Windows.Forms.Button
		Private WithEvents btnExportPdf As System.Windows.Forms.Button
		Private saveFileDialogXls As System.Windows.Forms.SaveFileDialog
		Private saveFileDialogPdf As System.Windows.Forms.SaveFileDialog
		Private edCity As System.Windows.Forms.TextBox
		Private labelCity As System.Windows.Forms.Label
		Private label1 As System.Windows.Forms.Label
		Private WithEvents linkLabel1 As System.Windows.Forms.LinkLabel
		Private panel1 As Panel
		Private label2 As Label
	End Class
End Namespace

