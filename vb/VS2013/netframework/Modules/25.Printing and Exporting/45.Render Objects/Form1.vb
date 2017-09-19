Imports System.Drawing.Drawing2D
Imports System.Collections
Imports System.ComponentModel
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Render
Imports System.IO
Imports System.Reflection

Imports System.Text



Namespace RenderObjects
	''' <summary>
	''' An Example on how to render a chart.
	''' </summary>
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Public Sub New()
			InitializeComponent()
			ResizeToolbar(mainToolbar)
		End Sub

		Private Sub ResizeToolbar(ByVal toolbar As ToolStrip)

			Using gr As Graphics = CreateGraphics()
				Dim xFactor As Double = gr.DpiX / 96.0
				Dim yFactor As Double = gr.DpiY / 96.0
				toolbar.ImageScalingSize = New Size(CInt(Fix(24 * xFactor)), CInt(Fix(24 * yFactor)))
				toolbar.Width = 0 'force a recalc of the buttons.
			End Using
		End Sub


		#Region "Global variables"
		Private Xls As XlsFile
		Private ValueRange As TXlsNamedRange
		Private MinValue As Double
		Private MaxValue As Double
		Private StepValue As Double
		Private ActualValue As Double

		Private ChartIndex As Integer
		Private ChartProps As TShapeProperties
		#End Region


		Private Sub InitApp()
			Xls = New XlsFile()

			Dim TemplatePath As String = Path.Combine(Path.Combine(Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), ".."), ".."), "templates") & Path.DirectorySeparatorChar
			Dim di As New DirectoryInfo(TemplatePath)
			Dim fi() As FileInfo = di.GetFiles("*.xls")
			If fi.Length = 0 Then
				Throw New Exception("Sorry, no templates found in the templates folder.")
			End If

			cbTheme.Items.Clear()
			For Each f As FileInfo In fi
				cbTheme.Items.Add(New FileHolder(f.FullName))
			Next f

			cbTheme.SelectedIndex = 0
		End Sub

		Private Sub LoadFile(ByVal FileName As String)
			Xls.Open(FileName)

			ActualValue = 0

			ValueRange = Xls.GetNamedRange("Value", 0)
			If ValueRange Is Nothing Then
				Throw New Exception("There is no range named ""value"" in the template")
			End If

			MinValue = ReadDoubleName("Minimum")
			MaxValue = ReadDoubleName("Maximum")
			StepValue = ReadDoubleName("Step")

			ChartIndex = -1
			For i As Integer = 1 To Xls.ObjectCount
				Dim ObjName As String = Xls.GetObjectName(i)
				If String.Compare(ObjName, "DataChart", True) = 0 Then
					ChartIndex = i
					Exit For
				End If
			Next i

			If ChartIndex < 0 Then
				Throw New Exception("There is no object named ""DataChart"" in the template")
			End If
			ChartProps = Xls.GetObjectProperties(ChartIndex, True)
		End Sub

		Private Function ReadDoubleName(ByVal Name As String) As Double
			Dim Range As TXlsCellRange = Xls.GetNamedRange(Name, 0)
			If Range Is Nothing Then
				Throw New Exception("There is no range named " & Name & " in the template")
			End If

			Dim val As Object = Xls.GetCellValue(Range.Top, Range.Left)
			If Not(TypeOf val Is Double) Then
				Throw New Exception("The range named " & Name & " does not contain a number")
			End If
			Return CDbl(val)
		End Function

		Private Sub button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.Click
			Close()
		End Sub

		Private Sub updater_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles updater.Tick
			Try
				ActualValue += StepValue
				If ActualValue > MaxValue Then
					ActualValue = MinValue
				End If
				Xls.SetCellValue(ValueRange.Top, ValueRange.Left, ActualValue)
				Xls.Recalc()

				If chartBox.Image IsNot Nothing Then
					chartBox.Image.Dispose()
				End If
				chartBox.Image = GetChart()
			Catch ex As Exception 'We don't want any dialog popping up every second.
				labelError.Text = ex.Message
				labelError.Dock = DockStyle.Fill
				panelError.Dock = DockStyle.Fill
				panelError.Visible = True
				updater.Enabled = False

			End Try
		End Sub

		Private Function GetChart() As Image
			'We could get the chart with the following command, 
			'but it would be fixed size. In this example we are going to be a little more complex.

			'Xls.RenderObject(ChartIndex);

			'A more complex way to retrieve the chart, to show how to use
			'all parameters in renderobject.

			Dim ImageDimensions As TUIRectangle
			Dim Origin As TPointF
			Dim SizePixels As TUISize

			'First calculate the chart dimensions without actually rendering it. This is fast.
			Xls.RenderObject(ChartIndex, 96, ChartProps, SmoothingMode.AntiAlias, InterpolationMode.HighQualityBicubic, True, False, Origin, ImageDimensions, SizePixels)

			Dim dpi As Double = 96 'default screen resolution
			If SizePixels.Height > 0 AndAlso SizePixels.Width > 0 Then
				Dim AspectX As Double = CDbl(chartBox.Width) / SizePixels.Width
				Dim AspectY As Double = CDbl(chartBox.Height) / SizePixels.Height

				Dim Aspect As Double = Math.Max(AspectX, AspectY)
				'Make the dpi adjust the screen resolution and the size of the form.
				dpi = CDbl(96 * Aspect)
				If dpi < 20 Then
					dpi = 20
				End If
				If dpi > 500 Then
					dpi = 500
				End If
			End If

			Return Xls.RenderObject(ChartIndex, dpi, ChartProps, SmoothingMode.AntiAlias, InterpolationMode.HighQualityBicubic, True, True, Origin, ImageDimensions, SizePixels)


		End Function

		Private Sub cbTheme_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbTheme.SelectedIndexChanged
			If cbTheme.SelectedItem Is Nothing Then
				Return
			End If
			LoadFile((TryCast(cbTheme.SelectedItem, FileHolder)).FullName)
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRun.Click
			If Xls Is Nothing Then
				InitApp()
			End If
			updater.Enabled = True
			btnRun.Enabled = False
			btnCancel.Enabled = True
		End Sub

		Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
			updater.Enabled = False
			btnRun.Enabled = True
			btnCancel.Enabled = False
			panelError.Visible = False
		End Sub

	End Class

	Friend Class FileHolder
		Friend FullName As String
		Private Caption As String

		Friend Sub New(ByVal aFullName As String)
			FullName = aFullName
			Caption = Path.GetFileNameWithoutExtension(aFullName)
		End Sub

		Public Overrides Function ToString() As String
			Return Caption
		End Function

	End Class
End Namespace
