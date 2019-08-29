Imports System.Collections
Imports System.ComponentModel
Imports FlexCel.Core
Imports System.Text
Namespace ExportHTML
	Partial Public Class Mailform
		Inherits System.Windows.Forms.Form


		Private WithEvents btnEmail As System.Windows.Forms.Button
		Private edOutServer As System.Windows.Forms.TextBox
		Private label2 As System.Windows.Forms.Label
		Private WithEvents edTo As System.Windows.Forms.TextBox
		Private label3 As System.Windows.Forms.Label
		Private WithEvents edFrom As System.Windows.Forms.TextBox
		Private label4 As System.Windows.Forms.Label
		Private edSubject As System.Windows.Forms.TextBox
		Private label5 As System.Windows.Forms.Label
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
			Dim resources As New System.ComponentModel.ComponentResourceManager(GetType(Mailform))
			Me.btnEmail = New System.Windows.Forms.Button()
			Me.edOutServer = New System.Windows.Forms.TextBox()
			Me.label2 = New System.Windows.Forms.Label()
			Me.edTo = New System.Windows.Forms.TextBox()
			Me.label3 = New System.Windows.Forms.Label()
			Me.edFrom = New System.Windows.Forms.TextBox()
			Me.label4 = New System.Windows.Forms.Label()
			Me.edSubject = New System.Windows.Forms.TextBox()
			Me.label5 = New System.Windows.Forms.Label()
			Me.SuspendLayout()
			' 
			' btnEmail
			' 
			Me.btnEmail.BackColor = System.Drawing.SystemColors.Control
			Me.btnEmail.Image = (CType(resources.GetObject("btnEmail.Image"), System.Drawing.Image))
			Me.btnEmail.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
			Me.btnEmail.Location = New System.Drawing.Point(200, 200)
			Me.btnEmail.Name = "btnEmail"
			Me.btnEmail.Size = New System.Drawing.Size(70, 30)
			Me.btnEmail.TabIndex = 4
			Me.btnEmail.Text = "e-mail!"
			Me.btnEmail.TextAlign = System.Drawing.ContentAlignment.MiddleRight
			Me.btnEmail.UseVisualStyleBackColor = False
'			Me.btnEmail.Click += New System.EventHandler(Me.btnEmail_Click)
			' 
			' edOutServer
			' 
			Me.edOutServer.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.edOutServer.Location = New System.Drawing.Point(136, 144)
			Me.edOutServer.Name = "edOutServer"
			Me.edOutServer.Size = New System.Drawing.Size(304, 20)
			Me.edOutServer.TabIndex = 3
			Me.edOutServer.Text = "pop.mycompany.com"
			' 
			' label2
			' 
			Me.label2.Location = New System.Drawing.Point(8, 152)
			Me.label2.Name = "label2"
			Me.label2.Size = New System.Drawing.Size(128, 16)
			Me.label2.TabIndex = 10
			Me.label2.Text = "Outgoing Mail Server:"
			' 
			' edTo
			' 
			Me.edTo.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.edTo.Location = New System.Drawing.Point(136, 40)
			Me.edTo.Name = "edTo"
			Me.edTo.Size = New System.Drawing.Size(304, 20)
			Me.edTo.TabIndex = 1
			Me.edTo.Text = "user@hiscompany.com"
'			Me.edTo.Leave += New System.EventHandler(Me.edTo_Leave)
			' 
			' label3
			' 
			Me.label3.Location = New System.Drawing.Point(16, 40)
			Me.label3.Name = "label3"
			Me.label3.Size = New System.Drawing.Size(128, 16)
			Me.label3.TabIndex = 14
			Me.label3.Text = "To:"
			' 
			' edFrom
			' 
			Me.edFrom.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.edFrom.Location = New System.Drawing.Point(136, 8)
			Me.edFrom.Name = "edFrom"
			Me.edFrom.Size = New System.Drawing.Size(304, 20)
			Me.edFrom.TabIndex = 0
			Me.edFrom.Text = "myname@mycompany.com"
'			Me.edFrom.Leave += New System.EventHandler(Me.edFrom_Leave)
			' 
			' label4
			' 
			Me.label4.Location = New System.Drawing.Point(16, 16)
			Me.label4.Name = "label4"
			Me.label4.Size = New System.Drawing.Size(120, 16)
			Me.label4.TabIndex = 12
			Me.label4.Text = "From:"
			' 
			' edSubject
			' 
			Me.edSubject.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.edSubject.Location = New System.Drawing.Point(136, 80)
			Me.edSubject.Name = "edSubject"
			Me.edSubject.Size = New System.Drawing.Size(304, 20)
			Me.edSubject.TabIndex = 2
			Me.edSubject.Text = "A test from FlexCel"
			' 
			' label5
			' 
			Me.label5.Location = New System.Drawing.Point(16, 80)
			Me.label5.Name = "label5"
			Me.label5.Size = New System.Drawing.Size(128, 16)
			Me.label5.TabIndex = 16
			Me.label5.Text = "Subject:"
			' 
			' Mailform
			' 
			Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New System.Drawing.Size(472, 246)
			Me.Controls.Add(Me.edSubject)
			Me.Controls.Add(Me.edTo)
			Me.Controls.Add(Me.edFrom)
			Me.Controls.Add(Me.edOutServer)
			Me.Controls.Add(Me.label5)
			Me.Controls.Add(Me.label3)
			Me.Controls.Add(Me.label4)
			Me.Controls.Add(Me.label2)
			Me.Controls.Add(Me.btnEmail)
			Me.Name = "Mailform"
			Me.ShowInTaskbar = False
			Me.Text = "Send email..."
			Me.ResumeLayout(False)
			Me.PerformLayout()

		End Sub
		#End Region
	End Class
End Namespace

