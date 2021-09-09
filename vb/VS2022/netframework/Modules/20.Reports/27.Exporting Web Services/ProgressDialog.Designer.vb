Imports System.Collections
Imports System.ComponentModel
Imports System.Threading
Imports FlexCel.Render
Namespace ExportingWebServices
	Partial Public Class ProgressDialog
		Inherits System.Windows.Forms.Form

		Private statusBar1 As System.Windows.Forms.StatusBar
		Private statusBarPanel1 As System.Windows.Forms.StatusBarPanel
		Private statusBarPanelTime As System.Windows.Forms.StatusBarPanel
		Private labelCaption As System.Windows.Forms.Label
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
			Me.statusBar1 = New System.Windows.Forms.StatusBar()
			Me.statusBarPanel1 = New System.Windows.Forms.StatusBarPanel()
			Me.statusBarPanelTime = New System.Windows.Forms.StatusBarPanel()
			Me.labelCaption = New System.Windows.Forms.Label()
			Me.timer1 = New System.Timers.Timer()
			CType(Me.statusBarPanel1, System.ComponentModel.ISupportInitialize).BeginInit()
			CType(Me.statusBarPanelTime, System.ComponentModel.ISupportInitialize).BeginInit()
			CType(Me.timer1, System.ComponentModel.ISupportInitialize).BeginInit()
			Me.SuspendLayout()
			' 
			' statusBar1
			' 
			Me.statusBar1.Location = New System.Drawing.Point(0, 100)
			Me.statusBar1.Name = "statusBar1"
			Me.statusBar1.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() { Me.statusBarPanel1, Me.statusBarPanelTime})
			Me.statusBar1.ShowPanels = True
			Me.statusBar1.Size = New System.Drawing.Size(448, 22)
			Me.statusBar1.TabIndex = 1
			' 
			' statusBarPanel1
			' 
			Me.statusBarPanel1.BorderStyle = System.Windows.Forms.StatusBarPanelBorderStyle.None
			Me.statusBarPanel1.Text = "Elapsed Time:"
			Me.statusBarPanel1.Width = 80
			' 
			' statusBarPanelTime
			' 
			Me.statusBarPanelTime.BorderStyle = System.Windows.Forms.StatusBarPanelBorderStyle.None
			Me.statusBarPanelTime.Text = "0:00"
			' 
			' labelCaption
			' 
			Me.labelCaption.Anchor = (CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.labelCaption.Location = New System.Drawing.Point(16, 16)
			Me.labelCaption.Name = "labelCaption"
			Me.labelCaption.Size = New System.Drawing.Size(408, 16)
			Me.labelCaption.TabIndex = 2
			Me.labelCaption.Text = "Connecting to Webservice..."
			' 
			' timer1
			' 
			Me.timer1.Enabled = True
			Me.timer1.SynchronizingObject = Me
'			Me.timer1.Elapsed += New System.Timers.ElapsedEventHandler(Me.timer1_Elapsed)
			' 
			' ProgressDialog
			' 
			Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New System.Drawing.Size(448, 122)
			Me.Controls.Add(Me.labelCaption)
			Me.Controls.Add(Me.statusBar1)
			Me.Name = "ProgressDialog"
			Me.ShowInTaskbar = False
			Me.Text = "Please wait..."
'			Me.Closed += New System.EventHandler(Me.ProgressDialog_Closed)
			CType(Me.statusBarPanel1, System.ComponentModel.ISupportInitialize).EndInit()
			CType(Me.statusBarPanelTime, System.ComponentModel.ISupportInitialize).EndInit()
			CType(Me.timer1, System.ComponentModel.ISupportInitialize).EndInit()
			Me.ResumeLayout(False)

		End Sub
		#End Region
	End Class
End Namespace

