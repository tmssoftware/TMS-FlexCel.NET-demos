Imports System.Collections
Imports System.ComponentModel
Namespace GenericReports
	Partial Public Class EnterSQLDialog
		Inherits System.Windows.Forms.Form

		Private edSQL As System.Windows.Forms.TextBox
		Private panel1 As System.Windows.Forms.Panel
		Private label1 As System.Windows.Forms.Label
		Private panel2 As System.Windows.Forms.Panel
		Private btnCancel As System.Windows.Forms.Button
		Private button1 As System.Windows.Forms.Button
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
			Me.edSQL = New System.Windows.Forms.TextBox()
			Me.panel1 = New System.Windows.Forms.Panel()
			Me.label1 = New System.Windows.Forms.Label()
			Me.panel2 = New System.Windows.Forms.Panel()
			Me.btnCancel = New System.Windows.Forms.Button()
			Me.button1 = New System.Windows.Forms.Button()
			Me.panel1.SuspendLayout()
			Me.panel2.SuspendLayout()
			Me.SuspendLayout()
			' 
			' edSQL
			' 
			Me.edSQL.Dock = System.Windows.Forms.DockStyle.Fill
			Me.edSQL.Location = New System.Drawing.Point(0, 24)
			Me.edSQL.Multiline = True
			Me.edSQL.Name = "edSQL"
			Me.edSQL.Size = New System.Drawing.Size(520, 125)
			Me.edSQL.TabIndex = 0
			Me.edSQL.Text = "Select * from orders"
			' 
			' panel1
			' 
			Me.panel1.BackColor = System.Drawing.Color.Gray
			Me.panel1.Controls.Add(Me.label1)
			Me.panel1.Dock = System.Windows.Forms.DockStyle.Top
			Me.panel1.ForeColor = System.Drawing.Color.White
			Me.panel1.Location = New System.Drawing.Point(0, 0)
			Me.panel1.Name = "panel1"
			Me.panel1.Size = New System.Drawing.Size(520, 24)
			Me.panel1.TabIndex = 1
			' 
			' label1
			' 
			Me.label1.Font = New System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.label1.Location = New System.Drawing.Point(17, 6)
			Me.label1.Name = "label1"
			Me.label1.Size = New System.Drawing.Size(296, 23)
			Me.label1.TabIndex = 0
			Me.label1.Text = "Enter SQL to execute"
			' 
			' panel2
			' 
			Me.panel2.BackColor = System.Drawing.Color.Gray
			Me.panel2.Controls.Add(Me.btnCancel)
			Me.panel2.Controls.Add(Me.button1)
			Me.panel2.Dock = System.Windows.Forms.DockStyle.Bottom
			Me.panel2.ForeColor = System.Drawing.Color.White
			Me.panel2.Location = New System.Drawing.Point(0, 149)
			Me.panel2.Name = "panel2"
			Me.panel2.Size = New System.Drawing.Size(520, 40)
			Me.panel2.TabIndex = 2
			' 
			' btnCancel
			' 
			Me.btnCancel.Anchor = (CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.btnCancel.BackColor = System.Drawing.Color.FromArgb((CByte(192)), (CByte(0)), (CByte(0)))
			Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
			Me.btnCancel.ForeColor = System.Drawing.Color.White
			Me.btnCancel.Location = New System.Drawing.Point(400, 9)
			Me.btnCancel.Name = "btnCancel"
			Me.btnCancel.Size = New System.Drawing.Size(112, 23)
			Me.btnCancel.TabIndex = 5
			Me.btnCancel.Text = "Cancel"
			' 
			' button1
			' 
			Me.button1.Anchor = (CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.button1.BackColor = System.Drawing.Color.Green
			Me.button1.DialogResult = System.Windows.Forms.DialogResult.OK
			Me.button1.ForeColor = System.Drawing.Color.White
			Me.button1.Location = New System.Drawing.Point(280, 9)
			Me.button1.Name = "button1"
			Me.button1.Size = New System.Drawing.Size(112, 23)
			Me.button1.TabIndex = 4
			Me.button1.Text = "Ok"
			' 
			' EnterSQLDialog
			' 
			Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New System.Drawing.Size(520, 189)
			Me.Controls.Add(Me.edSQL)
			Me.Controls.Add(Me.panel2)
			Me.Controls.Add(Me.panel1)
			Me.Name = "EnterSQLDialog"
			Me.Text = "Information"
			Me.panel1.ResumeLayout(False)
			Me.panel2.ResumeLayout(False)
			Me.ResumeLayout(False)

		End Sub
		#End Region
	End Class
End Namespace

