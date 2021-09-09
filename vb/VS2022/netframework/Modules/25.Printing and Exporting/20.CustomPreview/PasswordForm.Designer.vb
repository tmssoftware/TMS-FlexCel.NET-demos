Imports System.ComponentModel
Namespace CustomPreview
	Partial Public Class PasswordForm
		Inherits System.Windows.Forms.Form

		Private label1 As System.Windows.Forms.Label
		Private btnOk As System.Windows.Forms.Button
		Private label3 As System.Windows.Forms.Label
		Private PasswordEdit As System.Windows.Forms.TextBox
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
			Me.btnOk = New System.Windows.Forms.Button()
			Me.label3 = New System.Windows.Forms.Label()
			Me.PasswordEdit = New System.Windows.Forms.TextBox()
			Me.SuspendLayout()
			' 
			' label1
			' 
			Me.label1.Location = New System.Drawing.Point(24, 16)
			Me.label1.Name = "label1"
			Me.label1.Size = New System.Drawing.Size(240, 23)
			Me.label1.TabIndex = 0
			Me.label1.Text = "Please enter the password to open this file:"
			' 
			' btnOk
			' 
			Me.btnOk.DialogResult = System.Windows.Forms.DialogResult.OK
			Me.btnOk.Location = New System.Drawing.Point(152, 112)
			Me.btnOk.Name = "btnOk"
			Me.btnOk.TabIndex = 1
			Me.btnOk.Text = "Ok"
			' 
			' label3
			' 
			Me.label3.Location = New System.Drawing.Point(40, 64)
			Me.label3.Name = "label3"
			Me.label3.Size = New System.Drawing.Size(64, 23)
			Me.label3.TabIndex = 5
			Me.label3.Text = "Password:"
			' 
			' PasswordEdit
			' 
			Me.PasswordEdit.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.PasswordEdit.Location = New System.Drawing.Point(112, 64)
			Me.PasswordEdit.Name = "PasswordEdit"
			Me.PasswordEdit.PasswordChar = "*"c
			Me.PasswordEdit.Size = New System.Drawing.Size(200, 20)
			Me.PasswordEdit.TabIndex = 0
			Me.PasswordEdit.Text = ""
			' 
			' PasswordForm
			' 
			Me.AcceptButton = Me.btnOk
			Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New System.Drawing.Size(408, 154)
			Me.Controls.Add(Me.label3)
			Me.Controls.Add(Me.PasswordEdit)
			Me.Controls.Add(Me.btnOk)
			Me.Controls.Add(Me.label1)
			Me.Name = "PasswordForm"
			Me.ShowInTaskbar = False
			Me.Text = "File is password protected."
			Me.ResumeLayout(False)

		End Sub
		#End Region
	End Class
End Namespace

