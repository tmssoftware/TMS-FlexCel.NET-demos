Imports System.Drawing.Drawing2D
Imports System.Collections
Imports System.ComponentModel
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Render
Imports System.IO
Imports System.Reflection
Imports System.Text
Namespace RenderObjects
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

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
			Me.panel1 = New System.Windows.Forms.Panel()
			Me.panelError = New System.Windows.Forms.Panel()
			Me.labelError = New System.Windows.Forms.Label()
			Me.chartBox = New System.Windows.Forms.PictureBox()
			Me.panel7 = New System.Windows.Forms.Panel()
			Me.cbTheme = New System.Windows.Forms.ComboBox()
			Me.label2 = New System.Windows.Forms.Label()
			Me.checkBox4 = New System.Windows.Forms.CheckBox()
			Me.updater = New System.Windows.Forms.Timer(Me.components)
			Me.mainToolbar = New System.Windows.Forms.ToolStrip()
			Me.btnRun = New System.Windows.Forms.ToolStripButton()
			Me.toolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
			Me.btnExit = New System.Windows.Forms.ToolStripButton()
			Me.btnCancel = New System.Windows.Forms.ToolStripButton()
			Me.panel1.SuspendLayout()
			Me.panelError.SuspendLayout()
			CType(Me.chartBox, System.ComponentModel.ISupportInitialize).BeginInit()
			Me.panel7.SuspendLayout()
			Me.mainToolbar.SuspendLayout()
			Me.SuspendLayout()
			' 
			' panel1
			' 
			Me.panel1.BackColor = System.Drawing.Color.White
			Me.panel1.Controls.Add(Me.panelError)
			Me.panel1.Controls.Add(Me.chartBox)
			Me.panel1.Controls.Add(Me.panel7)
			Me.panel1.Dock = System.Windows.Forms.DockStyle.Fill
			Me.panel1.Location = New System.Drawing.Point(0, 38)
			Me.panel1.Name = "panel1"
			Me.panel1.Size = New System.Drawing.Size(464, 392)
			Me.panel1.TabIndex = 3
			' 
			' panelError
			' 
			Me.panelError.Controls.Add(Me.labelError)
			Me.panelError.Location = New System.Drawing.Point(136, 128)
			Me.panelError.Name = "panelError"
			Me.panelError.Size = New System.Drawing.Size(200, 100)
			Me.panelError.TabIndex = 52
			Me.panelError.Visible = False
			' 
			' labelError
			' 
			Me.labelError.Location = New System.Drawing.Point(8, 16)
			Me.labelError.Name = "labelError"
			Me.labelError.Size = New System.Drawing.Size(100, 23)
			Me.labelError.TabIndex = 0
			' 
			' chartBox
			' 
			Me.chartBox.Anchor = (CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.chartBox.Location = New System.Drawing.Point(24, 120)
			Me.chartBox.Name = "chartBox"
			Me.chartBox.Size = New System.Drawing.Size(416, 250)
			Me.chartBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
			Me.chartBox.TabIndex = 51
			Me.chartBox.TabStop = False
			' 
			' panel7
			' 
			Me.panel7.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.panel7.BackColor = System.Drawing.Color.FromArgb((CInt((CByte(224)))), (CInt((CByte(224)))), (CInt((CByte(224)))))
			Me.panel7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.panel7.Controls.Add(Me.cbTheme)
			Me.panel7.Controls.Add(Me.label2)
			Me.panel7.Location = New System.Drawing.Point(16, 16)
			Me.panel7.Name = "panel7"
			Me.panel7.Size = New System.Drawing.Size(432, 72)
			Me.panel7.TabIndex = 44
			' 
			' cbTheme
			' 
			Me.cbTheme.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
			Me.cbTheme.Location = New System.Drawing.Point(8, 32)
			Me.cbTheme.Name = "cbTheme"
			Me.cbTheme.Size = New System.Drawing.Size(248, 21)
			Me.cbTheme.TabIndex = 46
'			Me.cbTheme.SelectedIndexChanged += New System.EventHandler(Me.cbTheme_SelectedIndexChanged)
			' 
			' label2
			' 
			Me.label2.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label2.Location = New System.Drawing.Point(8, 8)
			Me.label2.Name = "label2"
			Me.label2.Size = New System.Drawing.Size(192, 16)
			Me.label2.TabIndex = 19
			Me.label2.Text = "Select Theme:"
			' 
			' checkBox4
			' 
			Me.checkBox4.Location = New System.Drawing.Point(0, 0)
			Me.checkBox4.Name = "checkBox4"
			Me.checkBox4.Size = New System.Drawing.Size(104, 24)
			Me.checkBox4.TabIndex = 0
			' 
			' updater
			' 
'			Me.updater.Tick += New System.EventHandler(Me.updater_Tick)
			' 
			' mainToolbar
			' 
			Me.mainToolbar.Items.AddRange(New System.Windows.Forms.ToolStripItem() { Me.btnRun, Me.toolStripSeparator1, Me.btnExit, Me.btnCancel})
			Me.mainToolbar.Location = New System.Drawing.Point(0, 0)
			Me.mainToolbar.Name = "mainToolbar"
			Me.mainToolbar.Size = New System.Drawing.Size(464, 38)
			Me.mainToolbar.TabIndex = 11
			Me.mainToolbar.Text = "toolStrip1"
			' 
			' btnRun
			' 
			Me.btnRun.Image = (CType(resources.GetObject("btnRun.Image"), System.Drawing.Image))
			Me.btnRun.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnRun.Name = "btnRun"
			Me.btnRun.Size = New System.Drawing.Size(35, 35)
			Me.btnRun.Text = "Run!"
			Me.btnRun.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
'			Me.btnRun.Click += New System.EventHandler(Me.btnRun_Click)
			' 
			' toolStripSeparator1
			' 
			Me.toolStripSeparator1.Name = "toolStripSeparator1"
			Me.toolStripSeparator1.Size = New System.Drawing.Size(6, 38)
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
			' btnCancel
			' 
			Me.btnCancel.Enabled = False
			Me.btnCancel.Image = (CType(resources.GetObject("btnCancel.Image"), System.Drawing.Image))
			Me.btnCancel.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnCancel.Name = "btnCancel"
			Me.btnCancel.Size = New System.Drawing.Size(47, 35)
			Me.btnCancel.Text = "Cancel"
			Me.btnCancel.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
'			Me.btnCancel.Click += New System.EventHandler(Me.btnCancel_Click)
			' 
			' mainForm
			' 
			Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New System.Drawing.Size(464, 430)
			Me.Controls.Add(Me.panel1)
			Me.Controls.Add(Me.mainToolbar)
			Me.MaximumSize = New System.Drawing.Size(800, 800)
			Me.Name = "mainForm"
			Me.Text = "Using FlexCel to render just a part of a spreadshet"
			Me.panel1.ResumeLayout(False)
			Me.panelError.ResumeLayout(False)
			CType(Me.chartBox, System.ComponentModel.ISupportInitialize).EndInit()
			Me.panel7.ResumeLayout(False)
			Me.mainToolbar.ResumeLayout(False)
			Me.mainToolbar.PerformLayout()
			Me.ResumeLayout(False)
			Me.PerformLayout()

		End Sub
		#End Region

		Private panel1 As System.Windows.Forms.Panel
		Private checkBox4 As System.Windows.Forms.CheckBox
		Private panel7 As System.Windows.Forms.Panel
		Private label2 As System.Windows.Forms.Label
		Private chartBox As System.Windows.Forms.PictureBox
		Private WithEvents updater As System.Windows.Forms.Timer
		Private WithEvents cbTheme As System.Windows.Forms.ComboBox
		Private panelError As System.Windows.Forms.Panel
		Private labelError As System.Windows.Forms.Label
		Private mainToolbar As ToolStrip
		Private WithEvents btnRun As ToolStripButton
		Private toolStripSeparator1 As ToolStripSeparator
		Private WithEvents btnExit As ToolStripButton
		Private WithEvents btnCancel As ToolStripButton
	End Class
End Namespace

