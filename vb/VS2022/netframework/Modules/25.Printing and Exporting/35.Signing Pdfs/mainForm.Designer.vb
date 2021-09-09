Namespace SigningPdfs
	Partial Public Class mainForm
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
			Me.btnCreateAndSign = New System.Windows.Forms.Button()
			Me.cbVisibleSignature = New System.Windows.Forms.CheckBox()
			Me.OpenExcelDialog = New System.Windows.Forms.OpenFileDialog()
			Me.savePdfDialog = New System.Windows.Forms.SaveFileDialog()
			Me.SignaturePicture = New System.Windows.Forms.PictureBox()
			Me.OpenImageDialog = New System.Windows.Forms.OpenFileDialog()
			CType(Me.SignaturePicture, System.ComponentModel.ISupportInitialize).BeginInit()
			Me.SuspendLayout()
			' 
			' btnCreateAndSign
			' 
			Me.btnCreateAndSign.Image = My.Resources.acroread
			Me.btnCreateAndSign.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
			Me.btnCreateAndSign.Location = New System.Drawing.Point(24, 25)
			Me.btnCreateAndSign.Name = "btnCreateAndSign"
			Me.btnCreateAndSign.Size = New System.Drawing.Size(155, 30)
			Me.btnCreateAndSign.TabIndex = 0
			Me.btnCreateAndSign.Text = "Create and Sign Pdf"
			Me.btnCreateAndSign.UseVisualStyleBackColor = True
'			Me.btnCreateAndSign.Click += New System.EventHandler(Me.btnCreateAndSign_Click)
			' 
			' cbVisibleSignature
			' 
			Me.cbVisibleSignature.AutoSize = True
			Me.cbVisibleSignature.Location = New System.Drawing.Point(24, 78)
			Me.cbVisibleSignature.Name = "cbVisibleSignature"
			Me.cbVisibleSignature.Size = New System.Drawing.Size(167, 17)
			Me.cbVisibleSignature.TabIndex = 1
			Me.cbVisibleSignature.Text = "Visible Signature (in last page)"
			Me.cbVisibleSignature.UseVisualStyleBackColor = True
'			Me.cbVisibleSignature.CheckedChanged += New System.EventHandler(Me.cbVisibleSignature_CheckedChanged)
			' 
			' OpenExcelDialog
			' 
			Me.OpenExcelDialog.DefaultExt = "xls"
			Me.OpenExcelDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|Excel 97/2003|*.xls|Excel 2007|*.xlsx;*.xlsm|All files|*.*"
			Me.OpenExcelDialog.Title = "Select Excel file to convert..."
			' 
			' savePdfDialog
			' 
			Me.savePdfDialog.DefaultExt = "pdf"
			Me.savePdfDialog.Filter = "Pdf Files|*.pdf"
			Me.savePdfDialog.Title = "Select where to save the file..."
			' 
			' SignaturePicture
			' 
			Me.SignaturePicture.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.SignaturePicture.Image = My.Resources.sign
			Me.SignaturePicture.Location = New System.Drawing.Point(24, 110)
			Me.SignaturePicture.Name = "SignaturePicture"
			Me.SignaturePicture.Size = New System.Drawing.Size(155, 100)
			Me.SignaturePicture.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
			Me.SignaturePicture.TabIndex = 2
			Me.SignaturePicture.TabStop = False
'			Me.SignaturePicture.Click += New System.EventHandler(Me.SignaturePicture_Click)
			' 
			' OpenImageDialog
			' 
			Me.OpenImageDialog.Filter = "Supported Images|*.png;*.bmp*.jpg|All files|*.*"
			' 
			' mainForm
			' 
			Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New System.Drawing.Size(205, 100)
			Me.Controls.Add(Me.SignaturePicture)
			Me.Controls.Add(Me.cbVisibleSignature)
			Me.Controls.Add(Me.btnCreateAndSign)
			Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
			Me.Name = "mainForm"
			Me.Text = "Signing PDFs"
			CType(Me.SignaturePicture, System.ComponentModel.ISupportInitialize).EndInit()
			Me.ResumeLayout(False)
			Me.PerformLayout()

		End Sub

		#End Region

		Private WithEvents btnCreateAndSign As System.Windows.Forms.Button
		Private WithEvents cbVisibleSignature As System.Windows.Forms.CheckBox
		Private OpenExcelDialog As System.Windows.Forms.OpenFileDialog
		Private savePdfDialog As System.Windows.Forms.SaveFileDialog
		Private WithEvents SignaturePicture As System.Windows.Forms.PictureBox
		Private OpenImageDialog As System.Windows.Forms.OpenFileDialog
	End Class
End Namespace

