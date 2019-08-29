Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports System.Resources
Imports System.Globalization
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Report
Namespace DirectSQL
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Private WithEvents button1 As System.Windows.Forms.Button
		Private saveFileDialog1 As System.Windows.Forms.SaveFileDialog
		Private label1 As System.Windows.Forms.Label
		Private WithEvents btnCancel As System.Windows.Forms.Button
		Private startDate As System.Windows.Forms.DateTimePicker
		Private label2 As System.Windows.Forms.Label
		Private endDate As System.Windows.Forms.DateTimePicker
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
			Me.btnCancel = New System.Windows.Forms.Button()
			Me.startDate = New System.Windows.Forms.DateTimePicker()
			Me.label2 = New System.Windows.Forms.Label()
			Me.endDate = New System.Windows.Forms.DateTimePicker()
			Me.SuspendLayout()
			' 
			' button1
			' 
			Me.button1.Anchor = (CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.button1.BackColor = System.Drawing.Color.Green
			Me.button1.ForeColor = System.Drawing.Color.White
			Me.button1.Location = New System.Drawing.Point(104, 136)
			Me.button1.Name = "button1"
			Me.button1.Size = New System.Drawing.Size(112, 23)
			Me.button1.TabIndex = 0
			Me.button1.Text = "GO!"
			Me.button1.UseVisualStyleBackColor = False
'			Me.button1.Click += New System.EventHandler(Me.button1_Click)
			' 
			' saveFileDialog1
			' 
			Me.saveFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|Excel 97/2003|*.xls|Excel 2007|*.xlsx;*.xlsm|All files|*.*"
			Me.saveFileDialog1.RestoreDirectory = True
			' 
			' label1
			' 
			Me.label1.Location = New System.Drawing.Point(24, 24)
			Me.label1.Name = "label1"
			Me.label1.Size = New System.Drawing.Size(288, 24)
			Me.label1.TabIndex = 2
			Me.label1.Text = "A demo on how to use direct SQL on a report. "
			' 
			' btnCancel
			' 
			Me.btnCancel.Anchor = (CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.btnCancel.BackColor = System.Drawing.Color.FromArgb((CInt((CByte(192)))), (CInt((CByte(0)))), (CInt((CByte(0)))))
			Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
			Me.btnCancel.ForeColor = System.Drawing.Color.White
			Me.btnCancel.Location = New System.Drawing.Point(224, 136)
			Me.btnCancel.Name = "btnCancel"
			Me.btnCancel.Size = New System.Drawing.Size(112, 23)
			Me.btnCancel.TabIndex = 3
			Me.btnCancel.Text = "Cancel"
			Me.btnCancel.UseVisualStyleBackColor = False
'			Me.btnCancel.Click += New System.EventHandler(Me.btnCancel_Click)
			' 
			' startDate
			' 
			Me.startDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
			Me.startDate.Location = New System.Drawing.Point(24, 80)
			Me.startDate.Name = "startDate"
			Me.startDate.Size = New System.Drawing.Size(144, 20)
			Me.startDate.TabIndex = 4
			Me.startDate.Value = New Date(1996, 1, 1, 15, 55, 0, 0)
			' 
			' label2
			' 
			Me.label2.Location = New System.Drawing.Point(24, 56)
			Me.label2.Name = "label2"
			Me.label2.Size = New System.Drawing.Size(224, 16)
			Me.label2.TabIndex = 5
			Me.label2.Text = "Get orders between:"
			' 
			' endDate
			' 
			Me.endDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
			Me.endDate.Location = New System.Drawing.Point(192, 80)
			Me.endDate.Name = "endDate"
			Me.endDate.Size = New System.Drawing.Size(144, 20)
			Me.endDate.TabIndex = 6
			Me.endDate.Value = New Date(1997, 1, 1, 0, 0, 0, 0)
			' 
			' mainForm
			' 
			Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New System.Drawing.Size(368, 182)
			Me.Controls.Add(Me.endDate)
			Me.Controls.Add(Me.label2)
			Me.Controls.Add(Me.startDate)
			Me.Controls.Add(Me.btnCancel)
			Me.Controls.Add(Me.label1)
			Me.Controls.Add(Me.button1)
			Me.Name = "mainForm"
			Me.Text = "Direct SQL"
			Me.ResumeLayout(False)

		End Sub
		#End Region
	End Class
End Namespace

