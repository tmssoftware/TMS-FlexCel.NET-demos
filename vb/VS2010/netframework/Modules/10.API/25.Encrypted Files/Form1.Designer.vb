Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Namespace EncryptedFiles
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Private panel2 As System.Windows.Forms.Panel
		Private WithEvents btnExit As System.Windows.Forms.Button
		Private WithEvents btnGo As System.Windows.Forms.Button
		Private saveFileDialog1 As System.Windows.Forms.SaveFileDialog
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
			Me.panel2 = New System.Windows.Forms.Panel()
			Me.btnExit = New System.Windows.Forms.Button()
			Me.saveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
			Me.btnGo = New System.Windows.Forms.Button()
			Me.panel2.SuspendLayout()
			Me.SuspendLayout()
			' 
			' panel2
			' 
			Me.panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
			Me.panel2.Controls.Add(Me.btnExit)
			Me.panel2.Dock = System.Windows.Forms.DockStyle.Top
			Me.panel2.Location = New System.Drawing.Point(0, 0)
			Me.panel2.Name = "panel2"
			Me.panel2.Size = New System.Drawing.Size(336, 35)
			Me.panel2.TabIndex = 4
			' 
			' btnExit
			' 
			Me.btnExit.Anchor = (CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.btnExit.BackColor = System.Drawing.SystemColors.Control
			Me.btnExit.Image = My.Resources._4close
			Me.btnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
			Me.btnExit.Location = New System.Drawing.Point(272, 2)
			Me.btnExit.Name = "btnExit"
			Me.btnExit.Size = New System.Drawing.Size(56, 26)
			Me.btnExit.TabIndex = 2
			Me.btnExit.Text = "Exit"
			Me.btnExit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
			Me.btnExit.UseVisualStyleBackColor = False
'			Me.btnExit.Click += New System.EventHandler(Me.btnExit_Click)
			' 
			' saveFileDialog1
			' 
			Me.saveFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|Excel 97/2003|*.xls|Excel 2007|*.xlsx;*.xlsm|All " & "files|*.*"
			Me.saveFileDialog1.RestoreDirectory = True
			' 
			' btnGo
			' 
			Me.btnGo.BackColor = System.Drawing.SystemColors.Control
			Me.btnGo.Image = My.Resources._4gears
			Me.btnGo.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
			Me.btnGo.Location = New System.Drawing.Point(96, 72)
			Me.btnGo.Name = "btnGo"
			Me.btnGo.Size = New System.Drawing.Size(152, 30)
			Me.btnGo.TabIndex = 5
			Me.btnGo.Text = "Create Encrypted File"
			Me.btnGo.TextAlign = System.Drawing.ContentAlignment.MiddleRight
			Me.btnGo.UseVisualStyleBackColor = False
'			Me.btnGo.Click += New System.EventHandler(Me.btnGo_Click)
			' 
			' mainForm
			' 
			Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New System.Drawing.Size(336, 122)
			Me.Controls.Add(Me.btnGo)
			Me.Controls.Add(Me.panel2)
			Me.Name = "mainForm"
			Me.Text = "Encrypted Excel Files"
			Me.panel2.ResumeLayout(False)
			Me.ResumeLayout(False)

		End Sub
		#End Region
	End Class
End Namespace

