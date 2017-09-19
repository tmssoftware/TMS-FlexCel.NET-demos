Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Namespace FlexCelImageExplorer
	Partial Public Class TCompressForm
		Inherits System.Windows.Forms.Form

		Private label1 As System.Windows.Forms.Label
		Private edPercent As System.Windows.Forms.NumericUpDown
		Private cbPixelFormat As System.Windows.Forms.ComboBox
		Private WithEvents btnOk As System.Windows.Forms.Button
		Private WithEvents btnCancel As System.Windows.Forms.Button
		Private panel1 As System.Windows.Forms.Panel
		Private panel2 As System.Windows.Forms.Panel
		Private pictureBox1 As System.Windows.Forms.PictureBox
		Private cbTransparent As System.Windows.Forms.CheckBox
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
			Me.label1 = New System.Windows.Forms.Label()
			Me.edPercent = New System.Windows.Forms.NumericUpDown()
			Me.cbPixelFormat = New System.Windows.Forms.ComboBox()
			Me.btnOk = New System.Windows.Forms.Button()
			Me.btnCancel = New System.Windows.Forms.Button()
			Me.panel1 = New System.Windows.Forms.Panel()
			Me.panel2 = New System.Windows.Forms.Panel()
			Me.pictureBox1 = New System.Windows.Forms.PictureBox()
			Me.cbTransparent = New System.Windows.Forms.CheckBox()
			CType(Me.edPercent, System.ComponentModel.ISupportInitialize).BeginInit()
			Me.panel1.SuspendLayout()
			Me.panel2.SuspendLayout()
			Me.SuspendLayout()
			' 
			' label1
			' 
			Me.label1.Location = New System.Drawing.Point(16, 8)
			Me.label1.Name = "label1"
			Me.label1.Size = New System.Drawing.Size(144, 16)
			Me.label1.TabIndex = 0
			Me.label1.Text = "Change Resolution (%):"
			' 
			' edPercent
			' 
			Me.edPercent.Increment = New System.Decimal(New Integer() { 5, 0, 0, 0})
			Me.edPercent.Location = New System.Drawing.Point(176, 8)
			Me.edPercent.Minimum = New System.Decimal(New Integer() { 10, 0, 0, 0})
			Me.edPercent.Name = "edPercent"
			Me.edPercent.Size = New System.Drawing.Size(48, 20)
			Me.edPercent.TabIndex = 2
			Me.edPercent.Value = New System.Decimal(New Integer() { 60, 0, 0, 0})
			' 
			' cbPixelFormat
			' 
			Me.cbPixelFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
			Me.cbPixelFormat.Items.AddRange(New Object() { "1bpp (Black and White)", "8bpp (256 colors optimized palette)", "24bpp (true color)"})
			Me.cbPixelFormat.Location = New System.Drawing.Point(16, 40)
			Me.cbPixelFormat.Name = "cbPixelFormat"
			Me.cbPixelFormat.Size = New System.Drawing.Size(208, 21)
			Me.cbPixelFormat.TabIndex = 3
			' 
			' btnOk
			' 
			Me.btnOk.Anchor = System.Windows.Forms.AnchorStyles.Bottom
			Me.btnOk.FlatStyle = System.Windows.Forms.FlatStyle.System
			Me.btnOk.Location = New System.Drawing.Point(148, 168)
			Me.btnOk.Name = "btnOk"
			Me.btnOk.TabIndex = 4
			Me.btnOk.Text = "Ok"
'			Me.btnOk.Click += New System.EventHandler(Me.btnOk_Click)
			' 
			' btnCancel
			' 
			Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
			Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
			Me.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.System
			Me.btnCancel.Location = New System.Drawing.Point(244, 168)
			Me.btnCancel.Name = "btnCancel"
			Me.btnCancel.TabIndex = 5
			Me.btnCancel.Text = "Cancel"
'			Me.btnCancel.Click += New System.EventHandler(Me.btnCancel_Click)
			' 
			' panel1
			' 
			Me.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
			Me.panel1.Controls.Add(Me.cbTransparent)
			Me.panel1.Controls.Add(Me.edPercent)
			Me.panel1.Controls.Add(Me.cbPixelFormat)
			Me.panel1.Controls.Add(Me.label1)
			Me.panel1.Location = New System.Drawing.Point(16, 16)
			Me.panel1.Name = "panel1"
			Me.panel1.Size = New System.Drawing.Size(240, 136)
			Me.panel1.TabIndex = 6
			' 
			' panel2
			' 
			Me.panel2.Anchor = (CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
			Me.panel2.Controls.Add(Me.pictureBox1)
			Me.panel2.Location = New System.Drawing.Point(280, 16)
			Me.panel2.Name = "panel2"
			Me.panel2.Size = New System.Drawing.Size(168, 136)
			Me.panel2.TabIndex = 7
			' 
			' pictureBox1
			' 
			Me.pictureBox1.Dock = System.Windows.Forms.DockStyle.Fill
			Me.pictureBox1.Location = New System.Drawing.Point(0, 0)
			Me.pictureBox1.Name = "pictureBox1"
			Me.pictureBox1.Size = New System.Drawing.Size(164, 132)
			Me.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
			Me.pictureBox1.TabIndex = 0
			Me.pictureBox1.TabStop = False
			' 
			' cbTransparent
			' 
			Me.cbTransparent.Location = New System.Drawing.Point(16, 88)
			Me.cbTransparent.Name = "cbTransparent"
			Me.cbTransparent.TabIndex = 4
			Me.cbTransparent.Text = "Transparent"
			' 
			' TCompressForm
			' 
			Me.AcceptButton = Me.btnOk
			Me.CancelButton = Me.btnCancel
			Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New System.Drawing.Size(472, 214)
			Me.Controls.Add(Me.panel2)
			Me.Controls.Add(Me.panel1)
			Me.Controls.Add(Me.btnCancel)
			Me.Controls.Add(Me.btnOk)
			Me.Name = "TCompressForm"
			Me.Text = "Compression Options..."
'			Me.Load += New System.EventHandler(Me.TCompressForm_Load)
			CType(Me.edPercent, System.ComponentModel.ISupportInitialize).EndInit()
			Me.panel1.ResumeLayout(False)
			Me.panel2.ResumeLayout(False)
			Me.ResumeLayout(False)

		End Sub
		#End Region
	End Class
End Namespace

