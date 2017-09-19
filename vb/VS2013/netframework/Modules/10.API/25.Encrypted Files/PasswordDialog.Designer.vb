Imports System.Collections
Imports System.ComponentModel
Namespace EncryptedFiles
	Partial Public Class PasswordDialog
		Inherits System.Windows.Forms.Form

		Private label1 As System.Windows.Forms.Label
		Private label2 As System.Windows.Forms.Label
		Private PasswordEdit As System.Windows.Forms.TextBox
		Private label3 As System.Windows.Forms.Label
		Private btnOk As System.Windows.Forms.Button
		Public btnCancel As System.Windows.Forms.Button
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
			Me.label2 = New System.Windows.Forms.Label()
			Me.PasswordEdit = New System.Windows.Forms.TextBox()
			Me.label3 = New System.Windows.Forms.Label()
			Me.btnOk = New System.Windows.Forms.Button()
			Me.btnCancel = New System.Windows.Forms.Button()
			Me.SuspendLayout()
			' 
			' label1
			' 
			Me.label1.Location = New System.Drawing.Point(16, 8)
			Me.label1.Name = "label1"
			Me.label1.Size = New System.Drawing.Size(280, 23)
			Me.label1.TabIndex = 0
			Me.label1.Text = "Please enter the password to open the template."
			' 
			' label2
			' 
			Me.label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label2.Location = New System.Drawing.Point(16, 32)
			Me.label2.Name = "label2"
			Me.label2.Size = New System.Drawing.Size(176, 23)
			Me.label2.TabIndex = 1
			Me.label2.Text = "HINT: The password is 42"
			' 
			' PasswordEdit
			' 
			Me.PasswordEdit.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.PasswordEdit.Location = New System.Drawing.Point(128, 64)
			Me.PasswordEdit.Name = "PasswordEdit"
			Me.PasswordEdit.PasswordChar = "*"c
			Me.PasswordEdit.Size = New System.Drawing.Size(360, 20)
			Me.PasswordEdit.TabIndex = 0
			Me.PasswordEdit.Text = ""
			' 
			' label3
			' 
			Me.label3.Location = New System.Drawing.Point(56, 64)
			Me.label3.Name = "label3"
			Me.label3.Size = New System.Drawing.Size(64, 23)
			Me.label3.TabIndex = 3
			Me.label3.Text = "Password:"
			' 
			' btnOk
			' 
			Me.btnOk.DialogResult = System.Windows.Forms.DialogResult.OK
			Me.btnOk.Location = New System.Drawing.Point(128, 112)
			Me.btnOk.Name = "btnOk"
			Me.btnOk.TabIndex = 1
			Me.btnOk.Text = "Ok"
			' 
			' btnCancel
			' 
			Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
			Me.btnCancel.Location = New System.Drawing.Point(216, 112)
			Me.btnCancel.Name = "btnCancel"
			Me.btnCancel.TabIndex = 2
			Me.btnCancel.Text = "Cancel"
			' 
			' PasswordDialog
			' 
			Me.AcceptButton = Me.btnOk
			Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New System.Drawing.Size(520, 154)
			Me.Controls.Add(Me.btnCancel)
			Me.Controls.Add(Me.btnOk)
			Me.Controls.Add(Me.label3)
			Me.Controls.Add(Me.PasswordEdit)
			Me.Controls.Add(Me.label2)
			Me.Controls.Add(Me.label1)
			Me.Name = "PasswordDialog"
			Me.ShowInTaskbar = False
			Me.Text = "Information"
			Me.ResumeLayout(False)

		End Sub
		#End Region
	End Class
End Namespace

