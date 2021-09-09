Imports System.Collections
Imports System.ComponentModel
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports System.IO
Imports System.Reflection
Namespace GettingStarted
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
			Me.button1 = New System.Windows.Forms.Button()
			Me.saveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
			Me.label1 = New System.Windows.Forms.Label()
			Me.cbAutoOpen = New System.Windows.Forms.CheckBox()
			Me.SuspendLayout()
			' 
			' button1
			' 
			Me.button1.Anchor = System.Windows.Forms.AnchorStyles.Bottom
			Me.button1.Location = New System.Drawing.Point(78, 160)
			Me.button1.Name = "button1"
			Me.button1.Size = New System.Drawing.Size(75, 23)
			Me.button1.TabIndex = 0
			Me.button1.Text = "GO!"
'			Me.button1.Click += New System.EventHandler(Me.button1_Click)
			' 
			' saveFileDialog1
			' 
			Me.saveFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|Excel 97/2003|*.xls|Excel 2007|*.xlsx;*.xlsm|All " & "files|*.*"
			Me.saveFileDialog1.RestoreDirectory = True
			' 
			' label1
			' 
			Me.label1.Anchor = (CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.label1.BackColor = System.Drawing.Color.FromArgb((CInt((CByte(255)))), (CInt((CByte(255)))), (CInt((CByte(192)))))
			Me.label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.label1.Location = New System.Drawing.Point(16, 16)
			Me.label1.Name = "label1"
			Me.label1.Size = New System.Drawing.Size(205, 80)
			Me.label1.TabIndex = 1
			Me.label1.Text = "A first demo on how to create basic things with FlexCel API."
			' 
			' cbAutoOpen
			' 
			Me.cbAutoOpen.Location = New System.Drawing.Point(16, 104)
			Me.cbAutoOpen.Name = "cbAutoOpen"
			Me.cbAutoOpen.Size = New System.Drawing.Size(144, 48)
			Me.cbAutoOpen.TabIndex = 2
			Me.cbAutoOpen.Text = "Auto open the generated file without saving it"
			' 
			' mainForm
			' 
			Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New System.Drawing.Size(229, 197)
			Me.Controls.Add(Me.cbAutoOpen)
			Me.Controls.Add(Me.label1)
			Me.Controls.Add(Me.button1)
			Me.Name = "mainForm"
			Me.Text = "Getting Started"
			Me.ResumeLayout(False)

		End Sub
		#End Region

		Private WithEvents button1 As System.Windows.Forms.Button
		Private saveFileDialog1 As System.Windows.Forms.SaveFileDialog
		Private label1 As System.Windows.Forms.Label
		Private cbAutoOpen As System.Windows.Forms.CheckBox

	End Class
End Namespace

