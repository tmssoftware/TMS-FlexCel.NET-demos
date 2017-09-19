'On Borland Developer Studio you can not embed xls files directly.
'On this demo you can see how to do it, converting it to resx files with the
'XlsEmbed utility included on the Tools folder.

Imports System
Imports System.Drawing
Imports System.Collections
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Data
Imports System.IO
Imports System.Diagnostics
Imports System.Reflection
Imports System.Resources
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Report
Imports FlexCel.Demo.SharedData


Namespace TemplatesOnTheExe
	''' <summary>
	''' Summary description for Form1.
	''' </summary>
	Public Class mainForm
		Inherits System.Windows.Forms.Form

		Private WithEvents button1 As System.Windows.Forms.Button
		Private saveFileDialog1 As System.Windows.Forms.SaveFileDialog
		Private label1 As System.Windows.Forms.Label
		Private WithEvents btnCancel As System.Windows.Forms.Button
		Private WithEvents ordersReport As FlexCel.Report.FlexCelReport
		''' <summary>
		''' Required designer variable.
		''' </summary>
		Private components As System.ComponentModel.Container = Nothing

		Public Sub New()
			InitializeComponent()
			InitializeReports()
		End Sub

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
			Me.ordersReport = New FlexCel.Report.FlexCelReport()
			Me.label1 = New System.Windows.Forms.Label()
			Me.btnCancel = New System.Windows.Forms.Button()
			Me.SuspendLayout()
			' 
			' button1
			' 
			Me.button1.Anchor = (CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.button1.BackColor = System.Drawing.Color.Green
			Me.button1.ForeColor = System.Drawing.Color.White
			Me.button1.Location = New System.Drawing.Point(152, 88)
			Me.button1.Name = "button1"
			Me.button1.Size = New System.Drawing.Size(112, 23)
			Me.button1.TabIndex = 0
			Me.button1.Text = "GO!"
'			Me.button1.Click += New System.EventHandler(Me.button1_Click)
			' 
			' saveFileDialog1
			' 
			Me.saveFileDialog1.Filter = "Excel Files|*.xls"
			Me.saveFileDialog1.RestoreDirectory = True
			' 
			' ordersReport
			' 
			Me.ordersReport.Canceled = False
			Me.ordersReport.DeleteEmptyRanges = False
			Me.ordersReport.ErrorActions = FlexCel.Report.TErrorActions.None
'			Me.ordersReport.GetInclude += New FlexCel.Report.GetIncludeEventHandler(Me.ordersReport_GetInclude)
			' 
			' label1
			' 
			Me.label1.Location = New System.Drawing.Point(24, 24)
			Me.label1.Name = "label1"
			Me.label1.Size = New System.Drawing.Size(288, 24)
			Me.label1.TabIndex = 2
			Me.label1.Text = "How to read a template from the executable instead of an external file."
			' 
			' btnCancel
			' 
			Me.btnCancel.Anchor = (CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.btnCancel.BackColor = System.Drawing.Color.FromArgb((CByte(192)), (CByte(0)), (CByte(0)))
			Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
			Me.btnCancel.ForeColor = System.Drawing.Color.White
			Me.btnCancel.Location = New System.Drawing.Point(272, 88)
			Me.btnCancel.Name = "btnCancel"
			Me.btnCancel.Size = New System.Drawing.Size(112, 23)
			Me.btnCancel.TabIndex = 3
			Me.btnCancel.Text = "Cancel"
'			Me.btnCancel.Click += New System.EventHandler(Me.btnCancel_Click)
			' 
			' mainForm
			' 
			Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
			Me.ClientSize = New System.Drawing.Size(416, 133)
			Me.Controls.Add(Me.btnCancel)
			Me.Controls.Add(Me.label1)
			Me.Controls.Add(Me.button1)
			Me.Name = "mainForm"
			Me.Text = "Templates on the exe"
			Me.ResumeLayout(False)

		End Sub
		#End Region

		''' <summary>
		''' The main entry point for the application.
		''' </summary>
		<STAThread> _
		Shared Sub Main()
			Application.Run(New mainForm())
		End Sub

		''' <summary>
		''' The datamodule with the access tables.
		''' </summary>
		Private dataModule As SharedData

		''' <summary>
		''' This method will add the databases used by the report. As they won't change, we will do it only once.
		''' </summary>
		Private Sub InitializeReports()
			dataModule= New SharedData()
			ordersReport.AddTable(dataModule.nwind)
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles button1.Click
			AutoRun()
		End Sub

		Public Sub AutoRun()
			dataModule.LoadData()
			ordersReport.SetValue("Date", Date.Now)

			If saveFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
				Dim a As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()

				Dim rm As New System.Resources.ResourceManager("template", a)
				Dim WbData() As Byte=CType(rm.GetObject("Templates On The Exe.template.xls"), Byte())
				Using InStream As Stream = New MemoryStream(WbData)
					Using OutStream As New FileStream(saveFileDialog1.FileName, FileMode.Create)
						ordersReport.Run(InStream, OutStream)
					End Using
				End Using

				If MessageBox.Show("Do you want to open the generated file?","Confirm", MessageBoxButtons.YesNo)=System.Windows.Forms.DialogResult.Yes Then
					Process.Start(saveFileDialog1.FileName)
				End If
			End If

		End Sub

		''' <summary>
		''' This is the method that will be called by the ASP.NET front end. It returns an array of bytes 
		''' with the report data, so the ASP.NET application can stream it to the client.
		''' </summary>
		''' <returns>The generated file as a byte array.</returns>
		Public Function WebRun() As Byte()
			dataModule.LoadData()
			ordersReport.SetValue("Date", Date.Now)

			Dim DataPath As String= Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) &"\..\..\"

			Dim a As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
			Using OutStream As New MemoryStream()
				Dim rm As New System.Resources.ResourceManager("template", a)
				Dim WbData() As Byte=CType(rm.GetObject("Templates On The Exe.template.xls"), Byte())
				Using InStream As Stream = New MemoryStream(WbData)
					ordersReport.Run(InStream, OutStream)
					Return OutStream.ToArray()
				End Using
			End Using
		End Function


		Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
			Close()
		End Sub

		Private Sub ordersReport_GetInclude(ByVal sender As Object, ByVal e As FlexCel.Report.GetIncludeEventArgs) Handles ordersReport.GetInclude
			Dim a As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
			Dim rm As New System.Resources.ResourceManager("FileName), a)
			Dim Data() As Byte=CType(rm.GetObject(e.FileName), Byte())
			e.IncludeData=Data
		End Sub
	End Class

End Namespace
