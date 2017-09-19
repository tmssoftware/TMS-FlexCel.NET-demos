Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Report
Namespace GettingStartedReports
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Private WithEvents btnGo As System.Windows.Forms.Button
		Private saveFileDialog1 As System.Windows.Forms.SaveFileDialog

		Private edName As System.Windows.Forms.TextBox
		Private label1 As System.Windows.Forms.Label
		Private WithEvents btnCancel As System.Windows.Forms.Button
		Private label2 As System.Windows.Forms.Label
		Private edUrl As System.Windows.Forms.TextBox
		Private cbAutoOpen As System.Windows.Forms.CheckBox
		Private reportStart As FlexCel.Report.FlexCelReport

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
			Me.btnGo = New System.Windows.Forms.Button()
			Me.saveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
			Me.reportStart = New FlexCel.Report.FlexCelReport()
			Me.edName = New System.Windows.Forms.TextBox()
			Me.label1 = New System.Windows.Forms.Label()
			Me.btnCancel = New System.Windows.Forms.Button()
			Me.label2 = New System.Windows.Forms.Label()
			Me.edUrl = New System.Windows.Forms.TextBox()
			Me.cbAutoOpen = New System.Windows.Forms.CheckBox()
			Me.SuspendLayout()
			' 
			' btnGo
			' 
			Me.btnGo.Anchor = (CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.btnGo.BackColor = System.Drawing.Color.Green
			Me.btnGo.ForeColor = System.Drawing.Color.White
			Me.btnGo.Location = New System.Drawing.Point(152, 152)
			Me.btnGo.Name = "btnGo"
			Me.btnGo.Size = New System.Drawing.Size(112, 24)
			Me.btnGo.TabIndex = 0
			Me.btnGo.Text = "GO!"
			Me.btnGo.UseVisualStyleBackColor = False
'			Me.btnGo.Click += New System.EventHandler(Me.btnGo_Click)
			' 
			' saveFileDialog1
			' 
			Me.saveFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|Excel 97/2003|*.xls|Excel 2007|*.xlsx;*.xlsm|All " & "files|*.*"
			Me.saveFileDialog1.RestoreDirectory = True
			' 
			' reportStart
			' 
			Me.reportStart.AllowOverwritingFiles = True
			Me.reportStart.DeleteEmptyRanges = False
			' 
			' edName
			' 
			Me.edName.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.edName.Location = New System.Drawing.Point(24, 40)
			Me.edName.Name = "edName"
			Me.edName.Size = New System.Drawing.Size(360, 20)
			Me.edName.TabIndex = 1
			' 
			' label1
			' 
			Me.label1.Location = New System.Drawing.Point(24, 24)
			Me.label1.Name = "label1"
			Me.label1.Size = New System.Drawing.Size(160, 16)
			Me.label1.TabIndex = 2
			Me.label1.Text = "Tell me your name:"
			' 
			' btnCancel
			' 
			Me.btnCancel.Anchor = (CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.btnCancel.BackColor = System.Drawing.Color.FromArgb((CInt((CByte(192)))), (CInt((CByte(0)))), (CInt((CByte(0)))))
			Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
			Me.btnCancel.ForeColor = System.Drawing.Color.White
			Me.btnCancel.Location = New System.Drawing.Point(272, 152)
			Me.btnCancel.Name = "btnCancel"
			Me.btnCancel.Size = New System.Drawing.Size(112, 23)
			Me.btnCancel.TabIndex = 3
			Me.btnCancel.Text = "Cancel"
			Me.btnCancel.UseVisualStyleBackColor = False
'			Me.btnCancel.Click += New System.EventHandler(Me.btnCancel_Click)
			' 
			' label2
			' 
			Me.label2.Location = New System.Drawing.Point(28, 60)
			Me.label2.Name = "label2"
			Me.label2.Size = New System.Drawing.Size(228, 16)
			Me.label2.TabIndex = 5
			Me.label2.Text = "Your Home page (without http://)"
			' 
			' edUrl
			' 
			Me.edUrl.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.edUrl.Location = New System.Drawing.Point(28, 76)
			Me.edUrl.Name = "edUrl"
			Me.edUrl.Size = New System.Drawing.Size(360, 20)
			Me.edUrl.TabIndex = 4
			Me.edUrl.Text = "www.tmssoftware.com"
			' 
			' cbAutoOpen
			' 
			Me.cbAutoOpen.Location = New System.Drawing.Point(24, 104)
			Me.cbAutoOpen.Name = "cbAutoOpen"
			Me.cbAutoOpen.Size = New System.Drawing.Size(264, 24)
			Me.cbAutoOpen.TabIndex = 6
			Me.cbAutoOpen.Text = "Auto open the generated file without saving it"
			' 
			' mainForm
			' 
			Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New System.Drawing.Size(416, 190)
			Me.Controls.Add(Me.cbAutoOpen)
			Me.Controls.Add(Me.label2)
			Me.Controls.Add(Me.edUrl)
			Me.Controls.Add(Me.btnCancel)
			Me.Controls.Add(Me.label1)
			Me.Controls.Add(Me.edName)
			Me.Controls.Add(Me.btnGo)
			Me.Name = "mainForm"
			Me.Text = "Getting Started"
			Me.ResumeLayout(False)
			Me.PerformLayout()

		End Sub
		#End Region
	End Class
End Namespace

