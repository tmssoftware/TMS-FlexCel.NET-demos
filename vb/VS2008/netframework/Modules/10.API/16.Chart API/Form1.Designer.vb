Imports System.Collections
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports System.IO
Imports System.Reflection
Imports System.Text
Namespace ChartAPI
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Private WithEvents button1 As System.Windows.Forms.Button
		Private saveFileDialog1 As System.Windows.Forms.SaveFileDialog
		Private label1 As System.Windows.Forms.Label
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
			Me.button1 = New System.Windows.Forms.Button()
			Me.saveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
			Me.label1 = New System.Windows.Forms.Label()
			Me.SuspendLayout()
			' 
			' button1
			' 
			Me.button1.Anchor = System.Windows.Forms.AnchorStyles.Bottom
			Me.button1.Location = New System.Drawing.Point(136, 160)
			Me.button1.Name = "button1"
			Me.button1.Size = New System.Drawing.Size(75, 23)
			Me.button1.TabIndex = 0
			Me.button1.Text = "GO!"
'			Me.button1.Click += New System.EventHandler(Me.button1_Click)
			' 
			' saveFileDialog1
			' 
			Me.saveFileDialog1.DefaultExt = "xlsx"
			Me.saveFileDialog1.Filter = "Excel Files (2007 or newer)|*.xlsx"
			Me.saveFileDialog1.RestoreDirectory = True
			' 
			' label1
			' 
			Me.label1.Anchor = (CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.label1.BackColor = System.Drawing.Color.FromArgb((CInt((CByte(255)))), (CInt((CByte(255)))), (CInt((CByte(192)))))
			Me.label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.label1.Location = New System.Drawing.Point(16, 16)
			Me.label1.Name = "label1"
			Me.label1.Size = New System.Drawing.Size(320, 90)
			Me.label1.TabIndex = 1
			Me.label1.Text = "A demo showing how to create a simple chart with code. Use APIMate for more compl" & "ex charts"
			' 
			' mainForm
			' 
			Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New System.Drawing.Size(344, 197)
			Me.Controls.Add(Me.label1)
			Me.Controls.Add(Me.button1)
			Me.Name = "mainForm"
			Me.Text = "Chart API"
			Me.ResumeLayout(False)

		End Sub
		#End Region
	End Class
End Namespace

