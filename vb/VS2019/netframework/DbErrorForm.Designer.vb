Imports System.Collections
Imports System.ComponentModel
Namespace MainDemo
	Partial Public Class DbErrorForm
		Inherits System.Windows.Forms.Form

		Private label1 As System.Windows.Forms.Label
		Public edError As System.Windows.Forms.TextBox
		Private label2 As System.Windows.Forms.Label
		Private button1 As System.Windows.Forms.Button
		Private WithEvents LinkToDownload As System.Windows.Forms.LinkLabel
		Private WithEvents btnCopy As System.Windows.Forms.Button
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
			Me.edError = New System.Windows.Forms.TextBox()
			Me.label2 = New System.Windows.Forms.Label()
			Me.LinkToDownload = New System.Windows.Forms.LinkLabel()
			Me.button1 = New System.Windows.Forms.Button()
			Me.btnCopy = New System.Windows.Forms.Button()
			Me.label3 = New System.Windows.Forms.Label()
			Me.SuspendLayout()
			' 
			' label1
			' 
			Me.label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label1.Location = New System.Drawing.Point(16, 16)
			Me.label1.Name = "label1"
			Me.label1.Size = New System.Drawing.Size(416, 16)
			Me.label1.TabIndex = 0
			Me.label1.Text = "There has been an error connecting to the database:"
			' 
			' edError
			' 
			Me.edError.Anchor = (CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.edError.BackColor = System.Drawing.SystemColors.ControlDark
			Me.edError.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.edError.ForeColor = System.Drawing.SystemColors.ControlLightLight
			Me.edError.Location = New System.Drawing.Point(24, 40)
			Me.edError.Multiline = True
			Me.edError.Name = "edError"
			Me.edError.ReadOnly = True
			Me.edError.Size = New System.Drawing.Size(867, 80)
			Me.edError.TabIndex = 1
			' 
			' label2
			' 
			Me.label2.Anchor = (CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles))
			Me.label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label2.Location = New System.Drawing.Point(24, 136)
			Me.label2.Name = "label2"
			Me.label2.Size = New System.Drawing.Size(656, 23)
			Me.label2.TabIndex = 2
			Me.label2.Text = "Verify that you have SQL Server CE (Compact edition) installed. If you don't plea" & "se install it from:"
			' 
			' LinkToDownload
			' 
			Me.LinkToDownload.Anchor = (CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles))
			Me.LinkToDownload.Font = New System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.LinkToDownload.Location = New System.Drawing.Point(24, 159)
			Me.LinkToDownload.Name = "LinkToDownload"
			Me.LinkToDownload.Size = New System.Drawing.Size(688, 24)
			Me.LinkToDownload.TabIndex = 4
			Me.LinkToDownload.TabStop = True
			Me.LinkToDownload.Text = "http://www.microsoft.com/downloads/details.aspx?FamilyId=DC614AEE-7E1C-4881-9C32-" & "3A6CE53384D9&displaylang=en"
'			Me.LinkToDownload.LinkClicked += New System.Windows.Forms.LinkLabelLinkClickedEventHandler(Me.LinkToDownload_LinkClicked)
			' 
			' button1
			' 
			Me.button1.Anchor = System.Windows.Forms.AnchorStyles.Bottom
			Me.button1.DialogResult = System.Windows.Forms.DialogResult.OK
			Me.button1.Location = New System.Drawing.Point(437, 296)
			Me.button1.Name = "button1"
			Me.button1.Size = New System.Drawing.Size(75, 23)
			Me.button1.TabIndex = 7
			Me.button1.Text = "Ok"
			' 
			' btnCopy
			' 
			Me.btnCopy.Anchor = (CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.btnCopy.Location = New System.Drawing.Point(779, 156)
			Me.btnCopy.Name = "btnCopy"
			Me.btnCopy.Size = New System.Drawing.Size(112, 23)
			Me.btnCopy.TabIndex = 8
			Me.btnCopy.Text = "Copy To Clipboard"
'			Me.btnCopy.Click += New System.EventHandler(Me.btnCopy2_Click)
			' 
			' label3
			' 
			Me.label3.AutoSize = True
			Me.label3.Location = New System.Drawing.Point(24, 195)
			Me.label3.Name = "label3"
			Me.label3.Size = New System.Drawing.Size(797, 13)
			Me.label3.TabIndex = 9
			Me.label3.Text = "SQL Server CE installs with the default setup in Visual Studio 2008. It is a ligh" & "t database (2MB), that doesn't install any services or background processes in y" & "our system."
			' 
			' DbErrorForm
			' 
			Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New System.Drawing.Size(915, 333)
			Me.Controls.Add(Me.label3)
			Me.Controls.Add(Me.btnCopy)
			Me.Controls.Add(Me.button1)
			Me.Controls.Add(Me.LinkToDownload)
			Me.Controls.Add(Me.label2)
			Me.Controls.Add(Me.edError)
			Me.Controls.Add(Me.label1)
			Me.Name = "DbErrorForm"
			Me.ShowIcon = False
			Me.ShowInTaskbar = False
			Me.Text = "Database error"
			Me.ResumeLayout(False)
			Me.PerformLayout()

		End Sub
		#End Region

		Private label3 As Label
	End Class
End Namespace

