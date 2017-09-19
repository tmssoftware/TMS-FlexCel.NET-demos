Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Report
Imports FlexCel.Demo.SharedData
Namespace EncryptionAndSubtotals
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Private WithEvents button1 As System.Windows.Forms.Button
		Private saveFileDialog1 As System.Windows.Forms.SaveFileDialog
		Private WithEvents btnCancel As System.Windows.Forms.Button
		Private panel1 As System.Windows.Forms.Panel
		Private label2 As System.Windows.Forms.Label
		Private OpenPassTemplate As System.Windows.Forms.TextBox
		Private label1 As System.Windows.Forms.Label
		Private panel2 As System.Windows.Forms.Panel
		Private label3 As System.Windows.Forms.Label
		Private OpenPassGenerated As System.Windows.Forms.TextBox
		Private label4 As System.Windows.Forms.Label
		Private label5 As System.Windows.Forms.Label
		Private ModifyPassGenerated As System.Windows.Forms.TextBox
		Private label6 As System.Windows.Forms.Label
		Private ProtectWorkbookPass As System.Windows.Forms.TextBox
		Private label7 As System.Windows.Forms.Label
		Private ProtectSheetPass As System.Windows.Forms.TextBox
		Private label8 As System.Windows.Forms.Label
		Private encryptionType As System.Windows.Forms.ComboBox
		Private label9 As System.Windows.Forms.Label
		Private ReservingUser As System.Windows.Forms.TextBox
		Private RecommendReadOnly As System.Windows.Forms.CheckBox
		Private ProtectWorkbook As System.Windows.Forms.CheckBox
		Private ProtectSheet As System.Windows.Forms.CheckBox
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
			Me.button1 = New System.Windows.Forms.Button()
			Me.saveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
			Me.btnCancel = New System.Windows.Forms.Button()
			Me.panel1 = New System.Windows.Forms.Panel()
			Me.label2 = New System.Windows.Forms.Label()
			Me.OpenPassTemplate = New System.Windows.Forms.TextBox()
			Me.label1 = New System.Windows.Forms.Label()
			Me.panel2 = New System.Windows.Forms.Panel()
			Me.ProtectSheet = New System.Windows.Forms.CheckBox()
			Me.ProtectWorkbook = New System.Windows.Forms.CheckBox()
			Me.RecommendReadOnly = New System.Windows.Forms.CheckBox()
			Me.label9 = New System.Windows.Forms.Label()
			Me.ReservingUser = New System.Windows.Forms.TextBox()
			Me.encryptionType = New System.Windows.Forms.ComboBox()
			Me.label8 = New System.Windows.Forms.Label()
			Me.ProtectSheetPass = New System.Windows.Forms.TextBox()
			Me.label7 = New System.Windows.Forms.Label()
			Me.label6 = New System.Windows.Forms.Label()
			Me.ProtectWorkbookPass = New System.Windows.Forms.TextBox()
			Me.label5 = New System.Windows.Forms.Label()
			Me.ModifyPassGenerated = New System.Windows.Forms.TextBox()
			Me.label3 = New System.Windows.Forms.Label()
			Me.OpenPassGenerated = New System.Windows.Forms.TextBox()
			Me.label4 = New System.Windows.Forms.Label()
			Me.panel1.SuspendLayout()
			Me.panel2.SuspendLayout()
			Me.SuspendLayout()
			' 
			' button1
			' 
			Me.button1.Anchor = (CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.button1.BackColor = System.Drawing.Color.Green
			Me.button1.ForeColor = System.Drawing.Color.White
			Me.button1.Location = New System.Drawing.Point(312, 414)
			Me.button1.Name = "button1"
			Me.button1.Size = New System.Drawing.Size(112, 23)
			Me.button1.TabIndex = 0
			Me.button1.Text = "GO!"
			Me.button1.UseVisualStyleBackColor = False
'			Me.button1.Click += New System.EventHandler(Me.button1_Click)
			' 
			' saveFileDialog1
			' 
			Me.saveFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|Excel 97/2003|*.xls|Excel 2007|*.xlsx;*.xlsm|All " & "files|*.*"
			Me.saveFileDialog1.RestoreDirectory = True
			' 
			' btnCancel
			' 
			Me.btnCancel.Anchor = (CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.btnCancel.BackColor = System.Drawing.Color.FromArgb((CInt((CByte(192)))), (CInt((CByte(0)))), (CInt((CByte(0)))))
			Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
			Me.btnCancel.ForeColor = System.Drawing.Color.White
			Me.btnCancel.Location = New System.Drawing.Point(432, 414)
			Me.btnCancel.Name = "btnCancel"
			Me.btnCancel.Size = New System.Drawing.Size(112, 23)
			Me.btnCancel.TabIndex = 3
			Me.btnCancel.Text = "Cancel"
			Me.btnCancel.UseVisualStyleBackColor = False
'			Me.btnCancel.Click += New System.EventHandler(Me.btnCancel_Click)
			' 
			' panel1
			' 
			Me.panel1.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.panel1.Controls.Add(Me.label2)
			Me.panel1.Controls.Add(Me.OpenPassTemplate)
			Me.panel1.Controls.Add(Me.label1)
			Me.panel1.Location = New System.Drawing.Point(24, 16)
			Me.panel1.Name = "panel1"
			Me.panel1.Size = New System.Drawing.Size(496, 80)
			Me.panel1.TabIndex = 6
			' 
			' label2
			' 
			Me.label2.Location = New System.Drawing.Point(8, 32)
			Me.label2.Name = "label2"
			Me.label2.Size = New System.Drawing.Size(200, 16)
			Me.label2.TabIndex = 8
			Me.label2.Text = "Password to open the template:"
			' 
			' OpenPassTemplate
			' 
			Me.OpenPassTemplate.Location = New System.Drawing.Point(8, 48)
			Me.OpenPassTemplate.Name = "OpenPassTemplate"
			Me.OpenPassTemplate.Size = New System.Drawing.Size(200, 20)
			Me.OpenPassTemplate.TabIndex = 7
			Me.OpenPassTemplate.Text = "flexcel"
			' 
			' label1
			' 
			Me.label1.Location = New System.Drawing.Point(8, 8)
			Me.label1.Name = "label1"
			Me.label1.Size = New System.Drawing.Size(416, 24)
			Me.label1.TabIndex = 6
			Me.label1.Text = "The template is protected with a password to open. On this demo, it is ""flexcel"""
			' 
			' panel2
			' 
			Me.panel2.Anchor = (CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.panel2.Controls.Add(Me.ProtectSheet)
			Me.panel2.Controls.Add(Me.ProtectWorkbook)
			Me.panel2.Controls.Add(Me.RecommendReadOnly)
			Me.panel2.Controls.Add(Me.label9)
			Me.panel2.Controls.Add(Me.ReservingUser)
			Me.panel2.Controls.Add(Me.encryptionType)
			Me.panel2.Controls.Add(Me.label8)
			Me.panel2.Controls.Add(Me.ProtectSheetPass)
			Me.panel2.Controls.Add(Me.label7)
			Me.panel2.Controls.Add(Me.label6)
			Me.panel2.Controls.Add(Me.ProtectWorkbookPass)
			Me.panel2.Controls.Add(Me.label5)
			Me.panel2.Controls.Add(Me.ModifyPassGenerated)
			Me.panel2.Controls.Add(Me.label3)
			Me.panel2.Controls.Add(Me.OpenPassGenerated)
			Me.panel2.Controls.Add(Me.label4)
			Me.panel2.Location = New System.Drawing.Point(24, 112)
			Me.panel2.Name = "panel2"
			Me.panel2.Size = New System.Drawing.Size(496, 296)
			Me.panel2.TabIndex = 7
			' 
			' ProtectSheet
			' 
			Me.ProtectSheet.Checked = True
			Me.ProtectSheet.CheckState = System.Windows.Forms.CheckState.Checked
			Me.ProtectSheet.Location = New System.Drawing.Point(368, 203)
			Me.ProtectSheet.Name = "ProtectSheet"
			Me.ProtectSheet.Size = New System.Drawing.Size(64, 16)
			Me.ProtectSheet.TabIndex = 22
			Me.ProtectSheet.Text = "Protect"
			' 
			' ProtectWorkbook
			' 
			Me.ProtectWorkbook.Checked = True
			Me.ProtectWorkbook.CheckState = System.Windows.Forms.CheckState.Checked
			Me.ProtectWorkbook.Location = New System.Drawing.Point(368, 147)
			Me.ProtectWorkbook.Name = "ProtectWorkbook"
			Me.ProtectWorkbook.Size = New System.Drawing.Size(64, 16)
			Me.ProtectWorkbook.TabIndex = 21
			Me.ProtectWorkbook.Text = "Protect"
			' 
			' RecommendReadOnly
			' 
			Me.RecommendReadOnly.Location = New System.Drawing.Point(240, 243)
			Me.RecommendReadOnly.Name = "RecommendReadOnly"
			Me.RecommendReadOnly.Size = New System.Drawing.Size(168, 24)
			Me.RecommendReadOnly.TabIndex = 20
			Me.RecommendReadOnly.Text = "Recommend read only"
			' 
			' label9
			' 
			Me.label9.Location = New System.Drawing.Point(8, 227)
			Me.label9.Name = "label9"
			Me.label9.Size = New System.Drawing.Size(200, 16)
			Me.label9.TabIndex = 19
			Me.label9.Text = "Reserving user: (for modify password)"
			' 
			' ReservingUser
			' 
			Me.ReservingUser.Location = New System.Drawing.Point(8, 243)
			Me.ReservingUser.Name = "ReservingUser"
			Me.ReservingUser.Size = New System.Drawing.Size(200, 20)
			Me.ReservingUser.TabIndex = 18
			Me.ReservingUser.Text = "Flexcel User"
			' 
			' encryptionType
			' 
			Me.encryptionType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
			Me.encryptionType.Items.AddRange(New Object() { "Default Excel 97/2000 Encryption", "Excel 95 XOR Encryption"})
			Me.encryptionType.Location = New System.Drawing.Point(8, 84)
			Me.encryptionType.Name = "encryptionType"
			Me.encryptionType.Size = New System.Drawing.Size(424, 21)
			Me.encryptionType.TabIndex = 17
			' 
			' label8
			' 
			Me.label8.Location = New System.Drawing.Point(8, 48)
			Me.label8.Name = "label8"
			Me.label8.Size = New System.Drawing.Size(488, 33)
			Me.label8.TabIndex = 16
			Me.label8.Text = "Encryption type for xls files (xlsx uses Agile encryption). Note that this is onl" & "y needed when saving, as the encryption type is autodetected when opening:"
			' 
			' ProtectSheetPass
			' 
			Me.ProtectSheetPass.Location = New System.Drawing.Point(232, 176)
			Me.ProtectSheetPass.Name = "ProtectSheetPass"
			Me.ProtectSheetPass.Size = New System.Drawing.Size(120, 20)
			Me.ProtectSheetPass.TabIndex = 15
			Me.ProtectSheetPass.Text = "sheet"
			' 
			' label7
			' 
			Me.label7.Location = New System.Drawing.Point(232, 179)
			Me.label7.Name = "label7"
			Me.label7.Size = New System.Drawing.Size(248, 16)
			Me.label7.TabIndex = 14
			Me.label7.Text = "Password to protect the generated sheets:"
			' 
			' label6
			' 
			Me.label6.Location = New System.Drawing.Point(232, 131)
			Me.label6.Name = "label6"
			Me.label6.Size = New System.Drawing.Size(248, 16)
			Me.label6.TabIndex = 12
			Me.label6.Text = "Password to protect the generated workbook:"
			' 
			' ProtectWorkbookPass
			' 
			Me.ProtectWorkbookPass.Location = New System.Drawing.Point(232, 147)
			Me.ProtectWorkbookPass.Name = "ProtectWorkbookPass"
			Me.ProtectWorkbookPass.Size = New System.Drawing.Size(120, 20)
			Me.ProtectWorkbookPass.TabIndex = 11
			Me.ProtectWorkbookPass.Text = "workbook"
			' 
			' label5
			' 
			Me.label5.Location = New System.Drawing.Point(8, 179)
			Me.label5.Name = "label5"
			Me.label5.Size = New System.Drawing.Size(200, 16)
			Me.label5.TabIndex = 10
			Me.label5.Text = "Password to modify the generated file:"
			' 
			' ModifyPassGenerated
			' 
			Me.ModifyPassGenerated.Location = New System.Drawing.Point(8, 195)
			Me.ModifyPassGenerated.Name = "ModifyPassGenerated"
			Me.ModifyPassGenerated.Size = New System.Drawing.Size(200, 20)
			Me.ModifyPassGenerated.TabIndex = 9
			Me.ModifyPassGenerated.Text = "modify"
			' 
			' label3
			' 
			Me.label3.Location = New System.Drawing.Point(8, 131)
			Me.label3.Name = "label3"
			Me.label3.Size = New System.Drawing.Size(200, 16)
			Me.label3.TabIndex = 8
			Me.label3.Text = "Password to open the generated file:"
			' 
			' OpenPassGenerated
			' 
			Me.OpenPassGenerated.Location = New System.Drawing.Point(8, 147)
			Me.OpenPassGenerated.Name = "OpenPassGenerated"
			Me.OpenPassGenerated.Size = New System.Drawing.Size(200, 20)
			Me.OpenPassGenerated.TabIndex = 7
			Me.OpenPassGenerated.Text = "open"
			' 
			' label4
			' 
			Me.label4.Location = New System.Drawing.Point(8, 8)
			Me.label4.Name = "label4"
			Me.label4.Size = New System.Drawing.Size(416, 40)
			Me.label4.TabIndex = 6
			Me.label4.Text = "Here we enter the passwords we want to protect the generated sheets and workbook." & " Leave them blank to have no password."
			' 
			' mainForm
			' 
			Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New System.Drawing.Size(552, 443)
			Me.Controls.Add(Me.panel2)
			Me.Controls.Add(Me.panel1)
			Me.Controls.Add(Me.btnCancel)
			Me.Controls.Add(Me.button1)
			Me.Name = "mainForm"
			Me.Text = "Encryption And Subtotals"
'			Me.Load += New System.EventHandler(Me.mainForm_Load)
			Me.panel1.ResumeLayout(False)
			Me.panel1.PerformLayout()
			Me.panel2.ResumeLayout(False)
			Me.panel2.PerformLayout()
			Me.ResumeLayout(False)

		End Sub
		#End Region
	End Class
End Namespace

