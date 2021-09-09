Imports System.Collections
Imports System.ComponentModel
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports System.IO
Imports System.Reflection
Imports System.Drawing.Drawing2D
Imports FlexCel.Pdf
Imports System.Runtime.InteropServices
Namespace PDFA
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Private panel1 As System.Windows.Forms.Panel
		Private exportDialog As System.Windows.Forms.SaveFileDialog
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
			Dim resources As New System.ComponentModel.ComponentResourceManager(GetType(mainForm))
			Me.panel1 = New System.Windows.Forms.Panel()
			Me.cbEmbed = New System.Windows.Forms.CheckBox()
			Me.cbPdfType = New System.Windows.Forms.ComboBox()
			Me.exportDialog = New System.Windows.Forms.SaveFileDialog()
			Me.toolStrip1 = New System.Windows.Forms.ToolStrip()
			Me.export = New System.Windows.Forms.ToolStripButton()
			Me.btnClose = New System.Windows.Forms.ToolStripButton()
			Me.panel1.SuspendLayout()
			Me.toolStrip1.SuspendLayout()
			Me.SuspendLayout()
			' 
			' panel1
			' 
			Me.panel1.BackColor = System.Drawing.Color.White
			Me.panel1.Controls.Add(Me.cbEmbed)
			Me.panel1.Controls.Add(Me.cbPdfType)
			Me.panel1.Dock = System.Windows.Forms.DockStyle.Fill
			Me.panel1.Location = New System.Drawing.Point(0, 38)
			Me.panel1.Name = "panel1"
			Me.panel1.Size = New System.Drawing.Size(372, 104)
			Me.panel1.TabIndex = 3
			' 
			' cbEmbed
			' 
			Me.cbEmbed.AutoSize = True
			Me.cbEmbed.Location = New System.Drawing.Point(12, 54)
			Me.cbEmbed.Name = "cbEmbed"
			Me.cbEmbed.Size = New System.Drawing.Size(351, 17)
			Me.cbEmbed.TabIndex = 5
			Me.cbEmbed.Text = "Embed xslx source file inside the PDF. (requires PDF/A3 or Standard)"
			Me.cbEmbed.UseVisualStyleBackColor = True
			' 
			' cbPdfType
			' 
			Me.cbPdfType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
			Me.cbPdfType.Items.AddRange(New Object() { "Standard", "PDF/A1a", "PDF/A1b", "PDF/A2a", "PDF/A2b", "PDF/A3a", "PDF/A3b"})
			Me.cbPdfType.Location = New System.Drawing.Point(12, 14)
			Me.cbPdfType.Name = "cbPdfType"
			Me.cbPdfType.Size = New System.Drawing.Size(144, 21)
			Me.cbPdfType.TabIndex = 35
			' 
			' exportDialog
			' 
			Me.exportDialog.DefaultExt = "pdf"
			Me.exportDialog.Filter = "Pdf files|*.pdf"
			' 
			' toolStrip1
			' 
			Me.toolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() { Me.export, Me.btnClose})
			Me.toolStrip1.Location = New System.Drawing.Point(0, 0)
			Me.toolStrip1.Name = "toolStrip1"
			Me.toolStrip1.Size = New System.Drawing.Size(372, 38)
			Me.toolStrip1.TabIndex = 4
			Me.toolStrip1.Text = "mainToolbar"
			' 
			' export
			' 
			Me.export.Image = (CType(resources.GetObject("export.Image"), System.Drawing.Image))
			Me.export.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.export.Name = "export"
			Me.export.Size = New System.Drawing.Size(69, 35)
			Me.export.Text = "Create PDF"
			Me.export.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
'			Me.export.Click += New System.EventHandler(Me.export_Click)
			' 
			' btnClose
			' 
			Me.btnClose.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
			Me.btnClose.Image = (CType(resources.GetObject("btnClose.Image"), System.Drawing.Image))
			Me.btnClose.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.btnClose.Name = "btnClose"
			Me.btnClose.Size = New System.Drawing.Size(59, 35)
			Me.btnClose.Text = "     E&xit     "
			Me.btnClose.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
'			Me.btnClose.Click += New System.EventHandler(Me.button2_Click)
			' 
			' mainForm
			' 
			Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New System.Drawing.Size(372, 142)
			Me.Controls.Add(Me.panel1)
			Me.Controls.Add(Me.toolStrip1)
			Me.Name = "mainForm"
			Me.Text = "PDF/A"
			Me.panel1.ResumeLayout(False)
			Me.panel1.PerformLayout()
			Me.toolStrip1.ResumeLayout(False)
			Me.toolStrip1.PerformLayout()
			Me.ResumeLayout(False)
			Me.PerformLayout()

		End Sub
		#End Region

		Private toolStrip1 As ToolStrip
		Private WithEvents export As ToolStripButton
		Private WithEvents btnClose As ToolStripButton
		Private cbEmbed As CheckBox
		Private cbPdfType As ComboBox
	End Class
End Namespace

