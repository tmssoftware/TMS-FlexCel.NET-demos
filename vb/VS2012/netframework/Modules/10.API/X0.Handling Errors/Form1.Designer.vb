Imports System.Collections
Imports System.ComponentModel
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports System.IO
Imports System.Reflection
Imports System.Text
Imports FlexCel.Render
Namespace HandlingErrors
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

			'Onhook the event handler. Since this is a form, we need to onhook the event when it is disposed or it would live forever.
			RemoveHandler FlexCelTrace.OnError, FlexCelTrace_OnErrorHandler

			MyBase.Dispose(disposing)
		End Sub

		#Region "Windows Form Designer generated code"
		''' <summary>
		''' Required method for Designer support - do not modify
		''' the contents of this method with the code editor.
		''' </summary>
		Private Sub InitializeComponent()
			Me.button1 = New System.Windows.Forms.Button()
			Me.saveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
			Me.label1 = New System.Windows.Forms.Label()
			Me.cbStopOnErrors = New System.Windows.Forms.CheckBox()
			Me.errorBox = New System.Windows.Forms.TextBox()
			Me.cbIgnoreFontErrors = New System.Windows.Forms.CheckBox()
			Me.SuspendLayout()
			' 
			' button1
			' 
			Me.button1.Anchor = System.Windows.Forms.AnchorStyles.Bottom
			Me.button1.Location = New System.Drawing.Point(244, 313)
			Me.button1.Name = "button1"
			Me.button1.TabIndex = 0
			Me.button1.Text = "GO!"
'			Me.button1.Click += New System.EventHandler(Me.button1_Click)
			' 
			' saveFileDialog1
			' 
			Me.saveFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|Excel 97/2003|*.xls|Excel 2007|*.xlsx;*.xlsm|All files|*.*"
			Me.saveFileDialog1.RestoreDirectory = True
			Me.saveFileDialog1.Title = "Save file as: (FILE WILL BE SAVED AS PDF TOO)"
			' 
			' label1
			' 
			Me.label1.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.label1.BackColor = System.Drawing.Color.FromArgb((CByte(255)), (CByte(255)), (CByte(192)))
			Me.label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.label1.Location = New System.Drawing.Point(16, 16)
			Me.label1.Name = "label1"
			Me.label1.Size = New System.Drawing.Size(536, 32)
			Me.label1.TabIndex = 1
			Me.label1.Text = "This demo shows how to handle non fatal errors in FlexCel by using the FlexCelTra" & "ce static class."
			' 
			' cbStopOnErrors
			' 
			Me.cbStopOnErrors.Location = New System.Drawing.Point(16, 64)
			Me.cbStopOnErrors.Name = "cbStopOnErrors"
			Me.cbStopOnErrors.Size = New System.Drawing.Size(400, 24)
			Me.cbStopOnErrors.TabIndex = 2
			Me.cbStopOnErrors.Text = "Stop on non fatal errors"
			' 
			' errorBox
			' 
			Me.errorBox.Anchor = (CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.errorBox.Font = New System.Drawing.Font("Arial Unicode MS", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, (CByte(0)))
			Me.errorBox.Location = New System.Drawing.Point(16, 128)
			Me.errorBox.Multiline = True
			Me.errorBox.Name = "errorBox"
			Me.errorBox.ReadOnly = True
			Me.errorBox.ScrollBars = System.Windows.Forms.ScrollBars.Both
			Me.errorBox.Size = New System.Drawing.Size(536, 160)
			Me.errorBox.TabIndex = 3
			Me.errorBox.Text = ""
			Me.errorBox.WordWrap = False
			' 
			' cbIgnoreFontErrors
			' 
			Me.cbIgnoreFontErrors.Location = New System.Drawing.Point(16, 88)
			Me.cbIgnoreFontErrors.Name = "cbIgnoreFontErrors"
			Me.cbIgnoreFontErrors.Size = New System.Drawing.Size(208, 24)
			Me.cbIgnoreFontErrors.TabIndex = 4
			Me.cbIgnoreFontErrors.Text = "Ignore font errors"
			' 
			' mainForm
			' 
			Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New System.Drawing.Size(560, 350)
			Me.Controls.Add(Me.cbIgnoreFontErrors)
			Me.Controls.Add(Me.errorBox)
			Me.Controls.Add(Me.cbStopOnErrors)
			Me.Controls.Add(Me.label1)
			Me.Controls.Add(Me.button1)
			Me.Name = "mainForm"
			Me.Text = "Handling non fatal errors."
			Me.ResumeLayout(False)

		End Sub
		#End Region

		Private WithEvents button1 As System.Windows.Forms.Button
		Private saveFileDialog1 As System.Windows.Forms.SaveFileDialog
		Private label1 As System.Windows.Forms.Label
		Private errorBox As System.Windows.Forms.TextBox
		Private cbStopOnErrors As System.Windows.Forms.CheckBox
		Private cbIgnoreFontErrors As System.Windows.Forms.CheckBox

	End Class
End Namespace

