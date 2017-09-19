Imports System.Collections
Imports System.ComponentModel
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports System.IO
Imports System.Reflection
Imports System.Globalization
Imports System.Drawing.Imaging
Imports System.Drawing.Drawing2D
Namespace RecalculationOfLinkedFiles
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Private panel2 As System.Windows.Forms.Panel
		Private WithEvents button2 As System.Windows.Forms.Button
		Private WithEvents CellA1 As System.Windows.Forms.TextBox
		Private panel1 As System.Windows.Forms.Panel
		Private label1 As System.Windows.Forms.Label
		Private label2 As System.Windows.Forms.Label
		Private label3 As System.Windows.Forms.Label
		Private Cell2 As System.Windows.Forms.TextBox
		Private label4 As System.Windows.Forms.Label
		Private label5 As System.Windows.Forms.Label
		Private label6 As System.Windows.Forms.Label
		Private label7 As System.Windows.Forms.Label
		Private Cell3 As System.Windows.Forms.TextBox
		Private label8 As System.Windows.Forms.Label
		Private label9 As System.Windows.Forms.Label
		Private Cell4 As System.Windows.Forms.TextBox
		Private label18 As System.Windows.Forms.Label
		Private panel3 As System.Windows.Forms.Panel
		Private WithEvents ChartA1 As System.Windows.Forms.TextBox
		Private WithEvents ChartB1 As System.Windows.Forms.TextBox
		Private WithEvents ChartB2 As System.Windows.Forms.TextBox
		Private WithEvents ChartA2 As System.Windows.Forms.TextBox
		Private WithEvents ChartB3 As System.Windows.Forms.TextBox
		Private WithEvents ChartA3 As System.Windows.Forms.TextBox
		Private chartBox As System.Windows.Forms.PictureBox
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
			Me.panel2 = New System.Windows.Forms.Panel()
			Me.button2 = New System.Windows.Forms.Button()
			Me.CellA1 = New System.Windows.Forms.TextBox()
			Me.panel1 = New System.Windows.Forms.Panel()
			Me.label8 = New System.Windows.Forms.Label()
			Me.label9 = New System.Windows.Forms.Label()
			Me.Cell4 = New System.Windows.Forms.TextBox()
			Me.label6 = New System.Windows.Forms.Label()
			Me.label7 = New System.Windows.Forms.Label()
			Me.Cell3 = New System.Windows.Forms.TextBox()
			Me.label5 = New System.Windows.Forms.Label()
			Me.label4 = New System.Windows.Forms.Label()
			Me.label3 = New System.Windows.Forms.Label()
			Me.Cell2 = New System.Windows.Forms.TextBox()
			Me.label2 = New System.Windows.Forms.Label()
			Me.label1 = New System.Windows.Forms.Label()
			Me.label18 = New System.Windows.Forms.Label()
			Me.ChartA1 = New System.Windows.Forms.TextBox()
			Me.panel3 = New System.Windows.Forms.Panel()
			Me.chartBox = New System.Windows.Forms.PictureBox()
			Me.ChartB3 = New System.Windows.Forms.TextBox()
			Me.ChartA3 = New System.Windows.Forms.TextBox()
			Me.ChartB2 = New System.Windows.Forms.TextBox()
			Me.ChartA2 = New System.Windows.Forms.TextBox()
			Me.ChartB1 = New System.Windows.Forms.TextBox()
			Me.panel2.SuspendLayout()
			Me.panel1.SuspendLayout()
			Me.panel3.SuspendLayout()
			CType(Me.chartBox, System.ComponentModel.ISupportInitialize).BeginInit()
			Me.SuspendLayout()
			' 
			' panel2
			' 
			Me.panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
			Me.panel2.Controls.Add(Me.button2)
			Me.panel2.Dock = System.Windows.Forms.DockStyle.Bottom
			Me.panel2.Location = New System.Drawing.Point(0, 422)
			Me.panel2.Name = "panel2"
			Me.panel2.Size = New System.Drawing.Size(760, 32)
			Me.panel2.TabIndex = 2
			' 
			' button2
			' 
			Me.button2.Anchor = (CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.button2.BackColor = System.Drawing.SystemColors.Control
			Me.button2.Image = (CType(resources.GetObject("button2.Image"), System.Drawing.Image))
			Me.button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
			Me.button2.Location = New System.Drawing.Point(697, 2)
			Me.button2.Name = "button2"
			Me.button2.Size = New System.Drawing.Size(56, 26)
			Me.button2.TabIndex = 2
			Me.button2.Text = "Exit"
			Me.button2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
			Me.button2.UseVisualStyleBackColor = False
'			Me.button2.Click += New System.EventHandler(Me.button2_Click)
			' 
			' CellA1
			' 
			Me.CellA1.Location = New System.Drawing.Point(24, 88)
			Me.CellA1.Name = "CellA1"
			Me.CellA1.Size = New System.Drawing.Size(100, 20)
			Me.CellA1.TabIndex = 3
'			Me.CellA1.TextChanged += New System.EventHandler(Me.CellA1_TextChanged)
			' 
			' panel1
			' 
			Me.panel1.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
			Me.panel1.Controls.Add(Me.label8)
			Me.panel1.Controls.Add(Me.label9)
			Me.panel1.Controls.Add(Me.Cell4)
			Me.panel1.Controls.Add(Me.label6)
			Me.panel1.Controls.Add(Me.label7)
			Me.panel1.Controls.Add(Me.Cell3)
			Me.panel1.Controls.Add(Me.label5)
			Me.panel1.Controls.Add(Me.label4)
			Me.panel1.Controls.Add(Me.label3)
			Me.panel1.Controls.Add(Me.Cell2)
			Me.panel1.Controls.Add(Me.label2)
			Me.panel1.Controls.Add(Me.label1)
			Me.panel1.Controls.Add(Me.CellA1)
			Me.panel1.Location = New System.Drawing.Point(8, 8)
			Me.panel1.Name = "panel1"
			Me.panel1.Size = New System.Drawing.Size(736, 136)
			Me.panel1.TabIndex = 4
			' 
			' label8
			' 
			Me.label8.Location = New System.Drawing.Point(520, 72)
			Me.label8.Name = "label8"
			Me.label8.Size = New System.Drawing.Size(160, 16)
			Me.label8.TabIndex = 15
			Me.label8.Text = "=[Third File.xls]Sheet1!A1 + 7"
			' 
			' label9
			' 
			Me.label9.Location = New System.Drawing.Point(520, 56)
			Me.label9.Name = "label9"
			Me.label9.Size = New System.Drawing.Size(100, 16)
			Me.label9.TabIndex = 14
			Me.label9.Text = "First File: A2"
			' 
			' Cell4
			' 
			Me.Cell4.Enabled = False
			Me.Cell4.Location = New System.Drawing.Point(520, 88)
			Me.Cell4.Name = "Cell4"
			Me.Cell4.Size = New System.Drawing.Size(152, 20)
			Me.Cell4.TabIndex = 13
			' 
			' label6
			' 
			Me.label6.Location = New System.Drawing.Point(328, 72)
			Me.label6.Name = "label6"
			Me.label6.Size = New System.Drawing.Size(184, 16)
			Me.label6.TabIndex = 12
			Me.label6.Text = "=[Second File.xls]Sheet1!A1 * 5"
			' 
			' label7
			' 
			Me.label7.Location = New System.Drawing.Point(328, 56)
			Me.label7.Name = "label7"
			Me.label7.Size = New System.Drawing.Size(100, 16)
			Me.label7.TabIndex = 11
			Me.label7.Text = "Third File: A1"
			' 
			' Cell3
			' 
			Me.Cell3.Enabled = False
			Me.Cell3.Location = New System.Drawing.Point(328, 88)
			Me.Cell3.Name = "Cell3"
			Me.Cell3.Size = New System.Drawing.Size(152, 20)
			Me.Cell3.TabIndex = 10
			' 
			' label5
			' 
			Me.label5.Location = New System.Drawing.Point(152, 72)
			Me.label5.Name = "label5"
			Me.label5.Size = New System.Drawing.Size(152, 16)
			Me.label5.TabIndex = 9
			Me.label5.Text = "=[First File.xls]Sheet1!A1 * 2"
			' 
			' label4
			' 
			Me.label4.Location = New System.Drawing.Point(24, 72)
			Me.label4.Name = "label4"
			Me.label4.Size = New System.Drawing.Size(100, 16)
			Me.label4.TabIndex = 8
			Me.label4.Text = "Constant"
			' 
			' label3
			' 
			Me.label3.Location = New System.Drawing.Point(152, 56)
			Me.label3.Name = "label3"
			Me.label3.Size = New System.Drawing.Size(100, 16)
			Me.label3.TabIndex = 7
			Me.label3.Text = "Second File: A1"
			' 
			' Cell2
			' 
			Me.Cell2.Enabled = False
			Me.Cell2.Location = New System.Drawing.Point(152, 88)
			Me.Cell2.Name = "Cell2"
			Me.Cell2.Size = New System.Drawing.Size(144, 20)
			Me.Cell2.TabIndex = 6
			' 
			' label2
			' 
			Me.label2.Location = New System.Drawing.Point(24, 56)
			Me.label2.Name = "label2"
			Me.label2.Size = New System.Drawing.Size(100, 16)
			Me.label2.TabIndex = 5
			Me.label2.Text = "First File: A1"
			' 
			' label1
			' 
			Me.label1.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.label1.Location = New System.Drawing.Point(16, 16)
			Me.label1.Name = "label1"
			Me.label1.Size = New System.Drawing.Size(704, 32)
			Me.label1.TabIndex = 4
			Me.label1.Text = "In this first example we will dynamically create 3 linked files. We will create a" & " workspace to link the files, and see how recalculation works."
			' 
			' label18
			' 
			Me.label18.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.label18.Location = New System.Drawing.Point(16, 16)
			Me.label18.Name = "label18"
			Me.label18.Size = New System.Drawing.Size(704, 32)
			Me.label18.TabIndex = 4
			Me.label18.Text = "This second example shows how to load files when we don't know a priori which fil" & "es we need to recalculate a file. To make it more interesting, we will use a cha" & "rt linked to other file."
			' 
			' ChartA1
			' 
			Me.ChartA1.Location = New System.Drawing.Point(16, 56)
			Me.ChartA1.Name = "ChartA1"
			Me.ChartA1.Size = New System.Drawing.Size(64, 20)
			Me.ChartA1.TabIndex = 3
			Me.ChartA1.Text = "1"
'			Me.ChartA1.TextChanged += New System.EventHandler(Me.Chart_TextChanged)
			' 
			' panel3
			' 
			Me.panel3.Anchor = (CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
			Me.panel3.Controls.Add(Me.chartBox)
			Me.panel3.Controls.Add(Me.ChartB3)
			Me.panel3.Controls.Add(Me.ChartA3)
			Me.panel3.Controls.Add(Me.ChartB2)
			Me.panel3.Controls.Add(Me.ChartA2)
			Me.panel3.Controls.Add(Me.ChartB1)
			Me.panel3.Controls.Add(Me.label18)
			Me.panel3.Controls.Add(Me.ChartA1)
			Me.panel3.Location = New System.Drawing.Point(8, 176)
			Me.panel3.Name = "panel3"
			Me.panel3.Size = New System.Drawing.Size(736, 225)
			Me.panel3.TabIndex = 5
			' 
			' chartBox
			' 
			Me.chartBox.Anchor = (CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.chartBox.Location = New System.Drawing.Point(176, 56)
			Me.chartBox.Name = "chartBox"
			Me.chartBox.Size = New System.Drawing.Size(544, 152)
			Me.chartBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
			Me.chartBox.TabIndex = 52
			Me.chartBox.TabStop = False
			' 
			' ChartB3
			' 
			Me.ChartB3.Location = New System.Drawing.Point(88, 104)
			Me.ChartB3.Name = "ChartB3"
			Me.ChartB3.Size = New System.Drawing.Size(64, 20)
			Me.ChartB3.TabIndex = 9
			Me.ChartB3.Text = "5"
'			Me.ChartB3.TextChanged += New System.EventHandler(Me.Chart_TextChanged)
			' 
			' ChartA3
			' 
			Me.ChartA3.Location = New System.Drawing.Point(16, 104)
			Me.ChartA3.Name = "ChartA3"
			Me.ChartA3.Size = New System.Drawing.Size(64, 20)
			Me.ChartA3.TabIndex = 8
			Me.ChartA3.Text = "3"
'			Me.ChartA3.TextChanged += New System.EventHandler(Me.Chart_TextChanged)
			' 
			' ChartB2
			' 
			Me.ChartB2.Location = New System.Drawing.Point(88, 80)
			Me.ChartB2.Name = "ChartB2"
			Me.ChartB2.Size = New System.Drawing.Size(64, 20)
			Me.ChartB2.TabIndex = 7
			Me.ChartB2.Text = "4"
'			Me.ChartB2.TextChanged += New System.EventHandler(Me.Chart_TextChanged)
			' 
			' ChartA2
			' 
			Me.ChartA2.Location = New System.Drawing.Point(16, 80)
			Me.ChartA2.Name = "ChartA2"
			Me.ChartA2.Size = New System.Drawing.Size(64, 20)
			Me.ChartA2.TabIndex = 6
			Me.ChartA2.Text = "2"
'			Me.ChartA2.TextChanged += New System.EventHandler(Me.Chart_TextChanged)
			' 
			' ChartB1
			' 
			Me.ChartB1.Location = New System.Drawing.Point(88, 56)
			Me.ChartB1.Name = "ChartB1"
			Me.ChartB1.Size = New System.Drawing.Size(64, 20)
			Me.ChartB1.TabIndex = 5
			Me.ChartB1.Text = "3"
'			Me.ChartB1.TextChanged += New System.EventHandler(Me.Chart_TextChanged)
			' 
			' mainForm
			' 
			Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New System.Drawing.Size(760, 454)
			Me.Controls.Add(Me.panel3)
			Me.Controls.Add(Me.panel1)
			Me.Controls.Add(Me.panel2)
			Me.Name = "mainForm"
			Me.Text = "Calculation of linked files"
			Me.panel2.ResumeLayout(False)
			Me.panel1.ResumeLayout(False)
			Me.panel1.PerformLayout()
			Me.panel3.ResumeLayout(False)
			Me.panel3.PerformLayout()
			CType(Me.chartBox, System.ComponentModel.ISupportInitialize).EndInit()
			Me.ResumeLayout(False)

		End Sub
		#End Region
	End Class
End Namespace

