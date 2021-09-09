Namespace VirtualMode
	Partial Public Class SheetSelectorForm
		''' <summary>
		''' Required designer variable.
		''' </summary>
		Private components As System.ComponentModel.IContainer = Nothing

		''' <summary>
		''' Clean up any resources being used.
		''' </summary>
		''' <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		Protected Overrides Sub Dispose(ByVal disposing As Boolean)
			If disposing AndAlso (components IsNot Nothing) Then
				components.Dispose()
			End If
			MyBase.Dispose(disposing)
		End Sub

		#Region "Windows Form Designer generated code"

		''' <summary>
		''' Required method for Designer support - do not modify
		''' the contents of this method with the code editor.
		''' </summary>
		Private Sub InitializeComponent()
			Me.panel1 = New System.Windows.Forms.Panel()
			Me.btnOk = New System.Windows.Forms.Button()
			Me.btnCancel = New System.Windows.Forms.Button()
			Me.SheetList = New System.Windows.Forms.ListBox()
			Me.panel1.SuspendLayout()
			Me.SuspendLayout()
			' 
			' panel1
			' 
			Me.panel1.Controls.Add(Me.btnCancel)
			Me.panel1.Controls.Add(Me.btnOk)
			Me.panel1.Dock = System.Windows.Forms.DockStyle.Bottom
			Me.panel1.Location = New System.Drawing.Point(0, 217)
			Me.panel1.Name = "panel1"
			Me.panel1.Size = New System.Drawing.Size(284, 45)
			Me.panel1.TabIndex = 0
			' 
			' btnOk
			' 
			Me.btnOk.DialogResult = System.Windows.Forms.DialogResult.OK
			Me.btnOk.Location = New System.Drawing.Point(116, 10)
			Me.btnOk.Name = "btnOk"
			Me.btnOk.Size = New System.Drawing.Size(75, 23)
			Me.btnOk.TabIndex = 0
			Me.btnOk.Text = "Ok"
			Me.btnOk.UseVisualStyleBackColor = True
			' 
			' btnCancel
			' 
			Me.btnCancel.Anchor = (CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
			Me.btnCancel.Location = New System.Drawing.Point(197, 10)
			Me.btnCancel.Name = "btnCancel"
			Me.btnCancel.Size = New System.Drawing.Size(75, 23)
			Me.btnCancel.TabIndex = 1
			Me.btnCancel.Text = "Cancel"
			Me.btnCancel.UseVisualStyleBackColor = True
			' 
			' SheetList
			' 
			Me.SheetList.Dock = System.Windows.Forms.DockStyle.Fill
			Me.SheetList.FormattingEnabled = True
			Me.SheetList.Location = New System.Drawing.Point(0, 0)
			Me.SheetList.Name = "SheetList"
			Me.SheetList.Size = New System.Drawing.Size(284, 217)
			Me.SheetList.TabIndex = 1
'			Me.SheetList.DoubleClick += New System.EventHandler(Me.SheetList_DoubleClick)
			' 
			' SheetSelectorForm
			' 
			Me.AcceptButton = Me.btnOk
			Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.CancelButton = Me.btnCancel
			Me.ClientSize = New System.Drawing.Size(284, 262)
			Me.Controls.Add(Me.SheetList)
			Me.Controls.Add(Me.panel1)
			Me.Name = "SheetSelectorForm"
			Me.Text = "Select sheet to load..."
			Me.panel1.ResumeLayout(False)
			Me.ResumeLayout(False)

		End Sub

		#End Region

		Private panel1 As System.Windows.Forms.Panel
		Private btnCancel As System.Windows.Forms.Button
		Private btnOk As System.Windows.Forms.Button
		Private WithEvents SheetList As System.Windows.Forms.ListBox
	End Class
End Namespace
