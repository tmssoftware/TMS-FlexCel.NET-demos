Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports System.Xml
Imports System.Net
Imports System.Threading
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Report
Imports FlexCel.Render
Namespace ExportingWebServices
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
            Me.btnExportXls = New System.Windows.Forms.Button()
            Me.saveFileDialogXls = New System.Windows.Forms.SaveFileDialog()
            Me.reportStart = New FlexCel.Report.FlexCelReport()
            Me.btnCancel = New System.Windows.Forms.Button()
            Me.label3 = New System.Windows.Forms.Label()
            Me.cbOffline = New System.Windows.Forms.CheckBox()
            Me.btnExportPdf = New System.Windows.Forms.Button()
            Me.saveFileDialogPdf = New System.Windows.Forms.SaveFileDialog()
            Me.edcity = New System.Windows.Forms.ComboBox()
            Me.SuspendLayout()
            ' 
            ' btnExportXls
            ' 
            Me.btnExportXls.Anchor = (CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
            Me.btnExportXls.BackColor = System.Drawing.Color.Green
            Me.btnExportXls.ForeColor = System.Drawing.Color.White
            Me.btnExportXls.Location = New System.Drawing.Point(16, 120)
            Me.btnExportXls.Name = "btnExportXls"
            Me.btnExportXls.Size = New System.Drawing.Size(112, 23)
            Me.btnExportXls.TabIndex = 0
            Me.btnExportXls.Text = "Export to Excel"
            Me.btnExportXls.UseVisualStyleBackColor = False
            '			Me.btnExportXls.Click += New System.EventHandler(Me.btnExportXls_Click)
            ' 
            ' saveFileDialogXls
            ' 
            Me.saveFileDialogXls.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|Excel 97/2003|*.xls|Excel 2007|*.xlsx;*.xlsm|All " & "files|*.*"
            Me.saveFileDialogXls.RestoreDirectory = True
            ' 
            ' reportStart
            ' 
            Me.reportStart.AllowOverwritingFiles = True
            Me.reportStart.DeleteEmptyRanges = False
            ' 
            ' btnCancel
            ' 
            Me.btnCancel.Anchor = (CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
            Me.btnCancel.BackColor = System.Drawing.Color.FromArgb((CInt((CByte(192)))), (CInt((CByte(0)))), (CInt((CByte(0)))))
            Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.btnCancel.ForeColor = System.Drawing.Color.White
            Me.btnCancel.Location = New System.Drawing.Point(272, 120)
            Me.btnCancel.Name = "btnCancel"
            Me.btnCancel.Size = New System.Drawing.Size(112, 23)
            Me.btnCancel.TabIndex = 3
            Me.btnCancel.Text = "Cancel"
            Me.btnCancel.UseVisualStyleBackColor = False
            '			Me.btnCancel.Click += New System.EventHandler(Me.btnCancel_Click)
            ' 
            ' label3
            ' 
            Me.label3.Location = New System.Drawing.Point(32, 19)
            Me.label3.Name = "label3"
            Me.label3.Size = New System.Drawing.Size(40, 85)
            Me.label3.TabIndex = 8
            Me.label3.Text = "City:"
            ' 
            ' cbOffline
            ' 
            Me.cbOffline.Checked = True
            Me.cbOffline.CheckState = System.Windows.Forms.CheckState.Checked
            Me.cbOffline.Location = New System.Drawing.Point(32, 80)
            Me.cbOffline.Name = "cbOffline"
            Me.cbOffline.Size = New System.Drawing.Size(352, 24)
            Me.cbOffline.TabIndex = 10
            Me.cbOffline.Text = "Use offline data. (do not actually connect to the web service)"
            ' 
            ' btnExportPdf
            ' 
            Me.btnExportPdf.Anchor = (CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
            Me.btnExportPdf.BackColor = System.Drawing.Color.SteelBlue
            Me.btnExportPdf.ForeColor = System.Drawing.Color.White
            Me.btnExportPdf.Location = New System.Drawing.Point(144, 120)
            Me.btnExportPdf.Name = "btnExportPdf"
            Me.btnExportPdf.Size = New System.Drawing.Size(112, 23)
            Me.btnExportPdf.TabIndex = 11
            Me.btnExportPdf.Text = "Export to Pdf"
            Me.btnExportPdf.UseVisualStyleBackColor = False
            '			Me.btnExportPdf.Click += New System.EventHandler(Me.btnExportPdf_Click)
            ' 
            ' saveFileDialogPdf
            ' 
            Me.saveFileDialogPdf.Filter = "Pdf Files|*.pdf"
            Me.saveFileDialogPdf.RestoreDirectory = True
            ' 
            ' edcity
            ' 
            Me.edcity.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
            Me.edcity.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
            Me.edcity.FormattingEnabled = True
            Me.edcity.Location = New System.Drawing.Point(78, 12)
            Me.edcity.MaxDropDownItems = 32
            Me.edcity.Name = "edcity"
            Me.edcity.Size = New System.Drawing.Size(306, 21)
            Me.edcity.TabIndex = 12
            '			Me.edcity.KeyDown += New System.Windows.Forms.KeyEventHandler(Me.edcity_KeyDown)
            ' 
            ' mainForm
            ' 
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0F, 13.0F)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(416, 157)
            Me.Controls.Add(Me.edcity)
            Me.Controls.Add(Me.btnExportPdf)
            Me.Controls.Add(Me.cbOffline)
            Me.Controls.Add(Me.label3)
            Me.Controls.Add(Me.btnCancel)
            Me.Controls.Add(Me.btnExportXls)
            Me.Name = "mainForm"
            Me.Text = "Exporting Web Services"
            Me.ResumeLayout(False)

        End Sub
#End Region

        Private WithEvents btnCancel As System.Windows.Forms.Button
        Private label3 As System.Windows.Forms.Label
        Private cbOffline As System.Windows.Forms.CheckBox
        Private WithEvents btnExportXls As System.Windows.Forms.Button
        Private WithEvents btnExportPdf As System.Windows.Forms.Button
        Private saveFileDialogXls As System.Windows.Forms.SaveFileDialog
        Private saveFileDialogPdf As System.Windows.Forms.SaveFileDialog
        Private reportStart As FlexCel.Report.FlexCelReport
        Private WithEvents edcity As ComboBox

    End Class
End Namespace


