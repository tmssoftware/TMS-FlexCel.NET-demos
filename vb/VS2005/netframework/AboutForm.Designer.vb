Imports System.Collections
Imports System.ComponentModel
Imports System.Reflection
Imports FlexCel.Core
Namespace MainDemo
	Partial Public Class AboutForm
		Inherits System.Windows.Forms.Form

		Private pictureBox1 As System.Windows.Forms.PictureBox
		Private label1 As System.Windows.Forms.Label
		Private WithEvents linkLabel1 As System.Windows.Forms.LinkLabel
		Private label2 As System.Windows.Forms.Label
		Private WithEvents button1 As System.Windows.Forms.Button
		Private WithEvents linkLabel2 As System.Windows.Forms.LinkLabel
		Private label3 As System.Windows.Forms.Label
		Private lblVersion As System.Windows.Forms.Label
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
			Dim resources As New System.ComponentModel.ComponentResourceManager(GetType(AboutForm))
			Me.pictureBox1 = New System.Windows.Forms.PictureBox()
			Me.label1 = New System.Windows.Forms.Label()
			Me.linkLabel1 = New System.Windows.Forms.LinkLabel()
			Me.label2 = New System.Windows.Forms.Label()
			Me.button1 = New System.Windows.Forms.Button()
			Me.linkLabel2 = New System.Windows.Forms.LinkLabel()
			Me.label3 = New System.Windows.Forms.Label()
			Me.lblVersion = New System.Windows.Forms.Label()
			CType(Me.pictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
			Me.SuspendLayout()
			' 
			' pictureBox1
			' 
			Me.pictureBox1.Anchor = (CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.pictureBox1.BackColor = System.Drawing.Color.Transparent
			Me.pictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
			Me.pictureBox1.Image = (CType(resources.GetObject("pictureBox1.Image"), System.Drawing.Image))
			Me.pictureBox1.Location = New System.Drawing.Point(392, 32)
			Me.pictureBox1.Name = "pictureBox1"
			Me.pictureBox1.Size = New System.Drawing.Size(48, 50)
			Me.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
			Me.pictureBox1.TabIndex = 0
			Me.pictureBox1.TabStop = False
			' 
			' label1
			' 
			Me.label1.BackColor = System.Drawing.Color.Transparent
			Me.label1.Font = New System.Drawing.Font("Garamond", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label1.Location = New System.Drawing.Point(40, 32)
			Me.label1.Name = "label1"
			Me.label1.Size = New System.Drawing.Size(384, 23)
			Me.label1.TabIndex = 1
			Me.label1.Text = "FLEXCEL WELL"
			' 
			' linkLabel1
			' 
			Me.linkLabel1.Anchor = (CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles))
			Me.linkLabel1.BackColor = System.Drawing.Color.Transparent
			Me.linkLabel1.Location = New System.Drawing.Point(88, 136)
			Me.linkLabel1.Name = "linkLabel1"
			Me.linkLabel1.Size = New System.Drawing.Size(176, 16)
			Me.linkLabel1.TabIndex = 2
			Me.linkLabel1.TabStop = True
			Me.linkLabel1.Text = "https://www.tmssoftware.com"
'			Me.linkLabel1.LinkClicked += New System.Windows.Forms.LinkLabelLinkClickedEventHandler(Me.linkLabel1_LinkClicked)
			' 
			' label2
			' 
			Me.label2.Anchor = (CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.label2.BackColor = System.Drawing.Color.Transparent
			Me.label2.Location = New System.Drawing.Point(40, 64)
			Me.label2.Name = "label2"
			Me.label2.Size = New System.Drawing.Size(400, 16)
			Me.label2.TabIndex = 3
			Me.label2.Text = "A repository with demos and documentation about flexcel."
			' 
			' button1
			' 
			Me.button1.BackColor = System.Drawing.Color.FromArgb((CInt((CByte(0)))), (CInt((CByte(192)))), (CInt((CByte(0)))))
			Me.button1.Cursor = System.Windows.Forms.Cursors.Default
			Me.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
			Me.button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.button1.ForeColor = System.Drawing.Color.White
			Me.button1.Location = New System.Drawing.Point(376, 176)
			Me.button1.Name = "button1"
			Me.button1.Size = New System.Drawing.Size(75, 23)
			Me.button1.TabIndex = 4
			Me.button1.Text = "OK"
			Me.button1.UseVisualStyleBackColor = False
'			Me.button1.Click += New System.EventHandler(Me.button1_Click)
			' 
			' linkLabel2
			' 
			Me.linkLabel2.Anchor = (CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles))
			Me.linkLabel2.BackColor = System.Drawing.Color.Transparent
			Me.linkLabel2.Location = New System.Drawing.Point(88, 160)
			Me.linkLabel2.Name = "linkLabel2"
			Me.linkLabel2.Size = New System.Drawing.Size(176, 16)
			Me.linkLabel2.TabIndex = 5
			Me.linkLabel2.TabStop = True
			Me.linkLabel2.Text = "mailto:help@tmssoftware.com"
'			Me.linkLabel2.LinkClicked += New System.Windows.Forms.LinkLabelLinkClickedEventHandler(Me.linkLabel2_LinkClicked)
			' 
			' label3
			' 
			Me.label3.Anchor = (CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.label3.BackColor = System.Drawing.Color.Transparent
			Me.label3.Font = New System.Drawing.Font("Times New Roman", 9.75F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label3.Location = New System.Drawing.Point(32, 112)
			Me.label3.Name = "label3"
			Me.label3.Size = New System.Drawing.Size(400, 18)
			Me.label3.TabIndex = 7
			Me.label3.Text = "For more information:"
			' 
			' lblVersion
			' 
			Me.lblVersion.Anchor = (CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.lblVersion.BackColor = System.Drawing.Color.Transparent
			Me.lblVersion.Location = New System.Drawing.Point(40, 88)
			Me.lblVersion.Name = "lblVersion"
			Me.lblVersion.Size = New System.Drawing.Size(400, 16)
			Me.lblVersion.TabIndex = 8
			' 
			' AboutForm
			' 
			Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.BackColor = System.Drawing.Color.Silver
			Me.BackgroundImage = (CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image))
			Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
			Me.ClientSize = New System.Drawing.Size(480, 237)
			Me.ControlBox = False
			Me.Controls.Add(Me.lblVersion)
			Me.Controls.Add(Me.label3)
			Me.Controls.Add(Me.linkLabel2)
			Me.Controls.Add(Me.pictureBox1)
			Me.Controls.Add(Me.label2)
			Me.Controls.Add(Me.button1)
			Me.Controls.Add(Me.linkLabel1)
			Me.Controls.Add(Me.label1)
			Me.DoubleBuffered = True
			Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
			Me.Name = "AboutForm"
			Me.ShowIcon = False
			Me.ShowInTaskbar = False
			Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
			Me.Text = "AboutForm"
			Me.TransparencyKey = System.Drawing.Color.Silver
'			Me.Load += New System.EventHandler(Me.AboutForm_Load)
			CType(Me.pictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
			Me.ResumeLayout(False)

		End Sub
		#End Region
	End Class
End Namespace

