Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports System.Data.OleDb
Imports System.Threading
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Report
Imports System.Xml
Namespace MetaTemplates
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Private saveFileDialog1 As System.Windows.Forms.SaveFileDialog
		Private panel2 As System.Windows.Forms.Panel
		Private WithEvents button2 As System.Windows.Forms.Button
		Private WithEvents btnExportExcel As System.Windows.Forms.Button
		Private cbFeeds As System.Windows.Forms.ComboBox
		Private cbOffline As System.Windows.Forms.CheckBox
		Private cbShowFeedCount As System.Windows.Forms.CheckBox
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
			Me.panel2 = New System.Windows.Forms.Panel()
			Me.button2 = New System.Windows.Forms.Button()
			Me.btnExportExcel = New System.Windows.Forms.Button()
			Me.cbFeeds = New System.Windows.Forms.ComboBox()
			Me.cbOffline = New System.Windows.Forms.CheckBox()
			Me.cbShowFeedCount = New System.Windows.Forms.CheckBox()
			Me.panel2.SuspendLayout()
			Me.SuspendLayout()
			' 
			' saveFileDialog1
			' 
			Me.saveFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|Excel 97/2003|*.xls|Excel 2007|*.xlsx;*.xlsm|All " & "files|*.*"
			Me.saveFileDialog1.RestoreDirectory = True
			' 
			' panel2
			' 
			Me.panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
			Me.panel2.Controls.Add(Me.button2)
			Me.panel2.Controls.Add(Me.btnExportExcel)
			Me.panel2.Controls.Add(Me.cbFeeds)
			Me.panel2.Dock = System.Windows.Forms.DockStyle.Top
			Me.panel2.Location = New System.Drawing.Point(0, 0)
			Me.panel2.Name = "panel2"
			Me.panel2.Size = New System.Drawing.Size(528, 40)
			Me.panel2.TabIndex = 3
			' 
			' button2
			' 
			Me.button2.Anchor = (CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.button2.BackColor = System.Drawing.SystemColors.Control
			Me.button2.Image = (CType(resources.GetObject("button2.Image"), System.Drawing.Image))
			Me.button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
			Me.button2.Location = New System.Drawing.Point(464, 2)
			Me.button2.Name = "button2"
			Me.button2.Size = New System.Drawing.Size(56, 26)
			Me.button2.TabIndex = 2
			Me.button2.Text = "Exit"
			Me.button2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
			Me.button2.UseVisualStyleBackColor = False
'			Me.button2.Click += New System.EventHandler(Me.button2_Click)
			' 
			' btnExportExcel
			' 
			Me.btnExportExcel.BackColor = System.Drawing.SystemColors.Control
			Me.btnExportExcel.Image = (CType(resources.GetObject("btnExportExcel.Image"), System.Drawing.Image))
			Me.btnExportExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
			Me.btnExportExcel.Location = New System.Drawing.Point(16, 2)
			Me.btnExportExcel.Name = "btnExportExcel"
			Me.btnExportExcel.Size = New System.Drawing.Size(120, 30)
			Me.btnExportExcel.TabIndex = 1
			Me.btnExportExcel.Text = "Export to Excel"
			Me.btnExportExcel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
			Me.btnExportExcel.UseVisualStyleBackColor = False
'			Me.btnExportExcel.Click += New System.EventHandler(Me.btnExportExcel_Click)
			' 
			' cbFeeds
			' 
			Me.cbFeeds.DisplayMember = "1"
			Me.cbFeeds.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
			Me.cbFeeds.Location = New System.Drawing.Point(200, 8)
			Me.cbFeeds.Name = "cbFeeds"
			Me.cbFeeds.Size = New System.Drawing.Size(216, 21)
			Me.cbFeeds.TabIndex = 4
			Me.cbFeeds.ValueMember = "1"
			' 
			' cbOffline
			' 
			Me.cbOffline.Checked = True
			Me.cbOffline.CheckState = System.Windows.Forms.CheckState.Checked
			Me.cbOffline.Location = New System.Drawing.Point(40, 56)
			Me.cbOffline.Name = "cbOffline"
			Me.cbOffline.Size = New System.Drawing.Size(368, 24)
			Me.cbOffline.TabIndex = 5
			Me.cbOffline.Text = "Use offline data (do not connect to internet)"
			' 
			' cbShowFeedCount
			' 
			Me.cbShowFeedCount.Location = New System.Drawing.Point(40, 80)
			Me.cbShowFeedCount.Name = "cbShowFeedCount"
			Me.cbShowFeedCount.Size = New System.Drawing.Size(360, 24)
			Me.cbShowFeedCount.TabIndex = 6
			Me.cbShowFeedCount.Text = "Show feed number column in the generated report."
			' 
			' mainForm
			' 
			Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New System.Drawing.Size(528, 126)
			Me.Controls.Add(Me.cbShowFeedCount)
			Me.Controls.Add(Me.cbOffline)
			Me.Controls.Add(Me.panel2)
			Me.Name = "mainForm"
			Me.Text = "Meta Templates"
'			Me.Load += New System.EventHandler(Me.mainForm_Load)
			Me.panel2.ResumeLayout(False)
			Me.ResumeLayout(False)

		End Sub
		#End Region
	End Class
End Namespace

