Imports System.Collections
Imports System.ComponentModel
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports System.IO
Imports System.Reflection
Namespace ConsolidatingFiles
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Private WithEvents button1 As System.Windows.Forms.Button
		Private saveFileDialog1 As System.Windows.Forms.SaveFileDialog
		Private label1 As System.Windows.Forms.Label
		Private openFileDialog1 As System.Windows.Forms.OpenFileDialog
		Private label2 As System.Windows.Forms.Label
		Private cbOnlyData As System.Windows.Forms.CheckBox
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
			Me.openFileDialog1 = New System.Windows.Forms.OpenFileDialog()
			Me.label2 = New System.Windows.Forms.Label()
			Me.cbOnlyData = New System.Windows.Forms.CheckBox()
			Me.SuspendLayout()
			' 
			' button1
			' 
			Me.button1.Anchor = System.Windows.Forms.AnchorStyles.Bottom
			Me.button1.Location = New System.Drawing.Point(180, 152)
			Me.button1.Name = "button1"
			Me.button1.TabIndex = 0
			Me.button1.Text = "GO!"
'			Me.button1.Click += New System.EventHandler(Me.button1_Click)
			' 
			' saveFileDialog1
			' 
			Me.saveFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|Excel 97/2003|*.xls|Excel 2007|*.xlsx;*.xlsm|All files|*.*"
			Me.saveFileDialog1.RestoreDirectory = True
			' 
			' label1
			' 
			Me.label1.Anchor = (CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.label1.BackColor = System.Drawing.Color.FromArgb((CByte(255)), (CByte(255)), (CByte(192)))
			Me.label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.label1.Location = New System.Drawing.Point(16, 16)
			Me.label1.Name = "label1"
			Me.label1.Size = New System.Drawing.Size(408, 32)
			Me.label1.TabIndex = 1
			Me.label1.Text = "A demo on how to consolidate several files into one."
			' 
			' openFileDialog1
			' 
			Me.openFileDialog1.DefaultExt = "xls"
			Me.openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|Excel 97/2003|*.xls|Excel 2007|*.xlsx;*.xlsm|All files|*.*"
			Me.openFileDialog1.Multiselect = True
			Me.openFileDialog1.Title = "Select ALL the files you want to consolidate."
			' 
			' label2
			' 
			Me.label2.Anchor = (CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.label2.BackColor = System.Drawing.Color.White
			Me.label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.label2.Location = New System.Drawing.Point(16, 64)
			Me.label2.Name = "label2"
			Me.label2.Size = New System.Drawing.Size(408, 32)
			Me.label2.TabIndex = 2
			Me.label2.Text = "After pressing the button, multi-select all the files you want with ctrl and shif" & "t."
			' 
			' cbOnlyData
			' 
			Me.cbOnlyData.Location = New System.Drawing.Point(24, 112)
			Me.cbOnlyData.Name = "cbOnlyData"
			Me.cbOnlyData.Size = New System.Drawing.Size(344, 24)
			Me.cbOnlyData.TabIndex = 3
			Me.cbOnlyData.Text = "Copy only data. (dont copy margins zoom, etc)"
			' 
			' mainForm
			' 
			Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New System.Drawing.Size(432, 197)
			Me.Controls.Add(Me.cbOnlyData)
			Me.Controls.Add(Me.label2)
			Me.Controls.Add(Me.label1)
			Me.Controls.Add(Me.button1)
			Me.Name = "mainForm"
			Me.Text = "Form1"
			Me.ResumeLayout(False)

		End Sub
		#End Region
	End Class
End Namespace

