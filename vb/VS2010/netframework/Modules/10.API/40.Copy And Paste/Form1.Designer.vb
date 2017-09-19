Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports System.Text
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Namespace CopyAndPaste
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
			Me.btnPaste = New System.Windows.Forms.Button()
			Me.btnNewFile = New System.Windows.Forms.Button()
			Me.btnCopy = New System.Windows.Forms.Button()
			Me.label1 = New System.Windows.Forms.Label()
			Me.label2 = New System.Windows.Forms.Label()
			Me.label3 = New System.Windows.Forms.Label()
			Me.label4 = New System.Windows.Forms.Label()
			Me.btnDragMe = New System.Windows.Forms.Button()
			Me.label5 = New System.Windows.Forms.Label()
			Me.DropHere = New System.Windows.Forms.Label()
			Me.btnOpenFile = New System.Windows.Forms.Button()
			Me.openFileDialog = New System.Windows.Forms.OpenFileDialog()
			Me.SuspendLayout()
			' 
			' btnPaste
			' 
			Me.btnPaste.Location = New System.Drawing.Point(44, 251)
			Me.btnPaste.Margin = New System.Windows.Forms.Padding(6, 6, 6, 6)
			Me.btnPaste.Name = "btnPaste"
			Me.btnPaste.Size = New System.Drawing.Size(138, 42)
			Me.btnPaste.TabIndex = 0
			Me.btnPaste.Text = "Paste"
'			Me.btnPaste.Click += New System.EventHandler(Me.btnPaste_Click)
			' 
			' btnNewFile
			' 
			Me.btnNewFile.Location = New System.Drawing.Point(44, 74)
			Me.btnNewFile.Margin = New System.Windows.Forms.Padding(6, 6, 6, 6)
			Me.btnNewFile.Name = "btnNewFile"
			Me.btnNewFile.Size = New System.Drawing.Size(138, 42)
			Me.btnNewFile.TabIndex = 1
			Me.btnNewFile.Text = "New File"
'			Me.btnNewFile.Click += New System.EventHandler(Me.btnNewFile_Click)
			' 
			' btnCopy
			' 
			Me.btnCopy.Location = New System.Drawing.Point(44, 428)
			Me.btnCopy.Margin = New System.Windows.Forms.Padding(6, 6, 6, 6)
			Me.btnCopy.Name = "btnCopy"
			Me.btnCopy.Size = New System.Drawing.Size(138, 42)
			Me.btnCopy.TabIndex = 2
			Me.btnCopy.Text = "Copy"
'			Me.btnCopy.Click += New System.EventHandler(Me.btnCopy_Click)
			' 
			' label1
			' 
			Me.label1.BackColor = System.Drawing.Color.LightSkyBlue
			Me.label1.Location = New System.Drawing.Point(44, 15)
			Me.label1.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
			Me.label1.Name = "label1"
			Me.label1.Size = New System.Drawing.Size(719, 44)
			Me.label1.TabIndex = 4
			Me.label1.Text = "1) Begin by creating a new file or opening an existing file..."
			Me.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
			' 
			' label2
			' 
			Me.label2.BackColor = System.Drawing.Color.LightSkyBlue
			Me.label2.Location = New System.Drawing.Point(44, 148)
			Me.label2.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
			Me.label2.Name = "label2"
			Me.label2.Size = New System.Drawing.Size(719, 44)
			Me.label2.TabIndex = 5
			Me.label2.Text = "2) Now go to Excel, copy some cells and paste them here..."
			Me.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
			' 
			' label3
			' 
			Me.label3.BackColor = System.Drawing.Color.LightSkyBlue
			Me.label3.Location = New System.Drawing.Point(44, 325)
			Me.label3.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
			Me.label3.Name = "label3"
			Me.label3.Size = New System.Drawing.Size(719, 44)
			Me.label3.TabIndex = 6
			Me.label3.Text = "3) After pasting, you can copy back the results to the clipboard"
			Me.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
			' 
			' label4
			' 
			Me.label4.BackColor = System.Drawing.Color.SteelBlue
			Me.label4.ForeColor = System.Drawing.Color.White
			Me.label4.Location = New System.Drawing.Point(44, 369)
			Me.label4.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
			Me.label4.Name = "label4"
			Me.label4.Size = New System.Drawing.Size(719, 44)
			Me.label4.TabIndex = 7
			Me.label4.Text = "Press the ""Copy"" button or drag the ""Drag Me!"" into Excel."
			Me.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
			' 
			' btnDragMe
			' 
			Me.btnDragMe.AllowDrop = True
			Me.btnDragMe.Location = New System.Drawing.Point(205, 428)
			Me.btnDragMe.Margin = New System.Windows.Forms.Padding(6, 6, 6, 6)
			Me.btnDragMe.Name = "btnDragMe"
			Me.btnDragMe.Size = New System.Drawing.Size(138, 42)
			Me.btnDragMe.TabIndex = 8
			Me.btnDragMe.Text = "Drag Me!"
'			Me.btnDragMe.MouseDown += New System.Windows.Forms.MouseEventHandler(Me.btnDragMe_MouseDown)
			' 
			' label5
			' 
			Me.label5.BackColor = System.Drawing.Color.SteelBlue
			Me.label5.ForeColor = System.Drawing.Color.White
			Me.label5.Location = New System.Drawing.Point(44, 192)
			Me.label5.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
			Me.label5.Name = "label5"
			Me.label5.Size = New System.Drawing.Size(719, 44)
			Me.label5.TabIndex = 10
			Me.label5.Text = "Press the ""Paste"" button or drag some cells from Excel into ""Drop Here!""."
			Me.label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
			' 
			' DropHere
			' 
			Me.DropHere.AllowDrop = True
			Me.DropHere.BackColor = System.Drawing.Color.FromArgb((CInt((CByte(192)))), (CInt((CByte(255)))), (CInt((CByte(192)))))
			Me.DropHere.Location = New System.Drawing.Point(200, 251)
			Me.DropHere.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
			Me.DropHere.Name = "DropHere"
			Me.DropHere.Size = New System.Drawing.Size(183, 42)
			Me.DropHere.TabIndex = 11
			Me.DropHere.Text = "Drop Here!"
			Me.DropHere.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'			Me.DropHere.DragDrop += New System.Windows.Forms.DragEventHandler(Me.DropHere_DragDrop)
'			Me.DropHere.DragOver += New System.Windows.Forms.DragEventHandler(Me.DropHere_DragOver)
			' 
			' btnOpenFile
			' 
			Me.btnOpenFile.Location = New System.Drawing.Point(205, 74)
			Me.btnOpenFile.Margin = New System.Windows.Forms.Padding(6)
			Me.btnOpenFile.Name = "btnOpenFile"
			Me.btnOpenFile.Size = New System.Drawing.Size(138, 42)
			Me.btnOpenFile.TabIndex = 12
			Me.btnOpenFile.Text = "Open File"
'			Me.btnOpenFile.Click += New System.EventHandler(Me.btnOpenFile_Click)
			' 
			' openFileDialog
			' 
			Me.openFileDialog.DefaultExt = "xls"
			Me.openFileDialog.Filter = "Excel Files|*.xls|All files|*.*"
			Me.openFileDialog.Title = "Select a file to preview"
			' 
			' mainForm
			' 
			Me.AutoScaleDimensions = New System.Drawing.SizeF(11F, 24F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New System.Drawing.Size(836, 615)
			Me.Controls.Add(Me.btnOpenFile)
			Me.Controls.Add(Me.DropHere)
			Me.Controls.Add(Me.label5)
			Me.Controls.Add(Me.btnDragMe)
			Me.Controls.Add(Me.label4)
			Me.Controls.Add(Me.label3)
			Me.Controls.Add(Me.label2)
			Me.Controls.Add(Me.label1)
			Me.Controls.Add(Me.btnCopy)
			Me.Controls.Add(Me.btnNewFile)
			Me.Controls.Add(Me.btnPaste)
			Me.Margin = New System.Windows.Forms.Padding(6, 6, 6, 6)
			Me.Name = "mainForm"
			Me.Text = "Copy and Paste Demo"
			Me.ResumeLayout(False)

		End Sub
		#End Region

		Private WithEvents btnPaste As Button
		Private WithEvents btnNewFile As Button
		Private WithEvents btnCopy As Button
		Private label1 As Label
		Private label2 As Label
		Private label3 As Label
		Private label4 As Label
		Private label5 As Label
		Private WithEvents btnDragMe As Button
		Private WithEvents DropHere As Label
		Private WithEvents btnOpenFile As Button
		Private openFileDialog As OpenFileDialog

	End Class
End Namespace

