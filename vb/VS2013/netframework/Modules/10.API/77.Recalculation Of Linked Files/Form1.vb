Imports System.Collections
Imports System.ComponentModel
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports System.IO
Imports System.Reflection
Imports System.Globalization
Imports System.Drawing.Imaging
Imports System.Drawing.Drawing2D

Namespace RecalculationOfLinkedFiles
	''' <summary>
	''' Shows how to recalculate linked files.
	''' </summary>
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles button2.Click
			Close()
		End Sub

		Private Sub CellA1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CellA1.TextChanged
			'This is a very slow way to do this (recreating the full 3 files each time you type a character)
			'but it is the best for what we want to show. (how to create and recalculate spreadsheets)
			'In a real world example you would keep the created files in memory and just recalculate them
			'when there is a change.
			CreateFilesAndRecalculate()
		End Sub

		''' <summary>
		''' This method will try to convert a text to a string, and if not possible, return the text.
		''' </summary>
		''' <param name="s"></param>
		''' <returns></returns>
		Private Function GetValue(ByVal s As String) As Object
			Dim d As Double
			If Double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, d) Then
				Return d
			End If
			Return s
		End Function

		Private Sub CreateFilesAndRecalculate()
			'Set up the files.
			Dim xls1 As New XlsFile()
			xls1.NewFile()

			xls1.SetCellValue(1, 1, GetValue(CellA1.Text))
			xls1.SetCellValue(2, 1, New TFormula("=[Third File.xls]Sheet1!A1 + 7"))

			Dim xls2 As New XlsFile()
			xls2.NewFile()
			xls2.SetCellValue(1, 1, New TFormula("=[First File.xls]Sheet1!A1 * 2"))

			Dim xls3 As New XlsFile()
			xls3.NewFile()
			xls3.SetCellValue(1, 1, New TFormula("=[Second File.xls]Sheet1!A1 * 5"))

			'Create a workspace to recalculate them.
			'In this case, as we know what files we need in advance, we will just add them to the workspace
			'For an example on how to load files on demand, take a look at the chart example in this demo.
			Dim Workspace As New TWorkspace()
			Workspace.Add("First File.xls", xls1)
			Workspace.Add("Second File.xls", xls2)
			Workspace.Add("Third File.xls", xls3)

			'Now that the workspace is set, we can recalculate. We could recalc() in the Workspace object or in any of the files in it.
			'The effect is the same, all files will be recalculated.
			'DO NOT RECALCULATE EVERY FILE. EACH TIME YOU CALCULATE ONE, YOU ARE CALCULATING THEM ALL.
			xls1.Recalc()

			'Ok, now it is time to show the results.
			Cell2.Text = Convert.ToString(CType(xls2.GetCellValue(1, 1), TFormula).Result)
			Cell3.Text = Convert.ToString(CType(xls3.GetCellValue(1, 1), TFormula).Result)
			Cell4.Text = Convert.ToString(CType(xls1.GetCellValue(2, 1), TFormula).Result)

			'In this example both the workspace and the xls files are local objects, so we don't need to worry about memory
			'If any of them is a global object, remember that keeping a reference to it will keep a reference to *ALL* the 
			'files in the workspace (even if you make Workspace = null). You might want to call Workspace.Clear() in that case before setting it to null.
		End Sub

		Private Sub Chart_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChartA1.TextChanged, ChartB3.TextChanged, ChartA3.TextChanged, ChartB2.TextChanged, ChartA2.TextChanged, ChartB1.TextChanged
			'Again, loading the file each time we press a key is incredibly silly. But for this example is ok,
			'since loading the files is what we actually want to show.

			Dim TemplatePath As String = Path.Combine(Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), ".."), "..")
			Dim xlsChart As New XlsFile()
			xlsChart.Open(Path.Combine(TemplatePath, "Chart.xls"))

			'Create a Workspace.
			'Note that if we didn't create this workspace, the chart would not show, since it wouldn't be able to
			'find the linked file. You can verify it by commenting the following lines
			Dim Workspace As New TWorkspace()
			Workspace.Add("Chart.xls", xlsChart) 'We always need to have the main file in the workspace.

			'The best thing here would be to also add "ChartData.xls" to the workspace, since we already know which file we need.
			'But since we already saw how to do that in the other example in this demo, we are going to pretend we don't know which files
			'we need, and load them on demand.
			'NOTE: DON'T LOAD FILES ON DEMAND UNLESS YOU REALLY NEED TO, SINCE YOU MIGHT BE CREATING A SECURITY RISK. Read the API GUIDE PDF for more information.
			AddHandler Workspace.LoadLinkedFile, AddressOf Workspace_LoadLinkedFile


			'Now that the Workspace is created, we can render the chart. We will use the code from "Render Objects" demo.
			If chartBox.Image IsNot Nothing Then
				chartBox.Image.Dispose()
			End If
			chartBox.Image = GetChart(xlsChart, 1) 'To do this well, we should name the chart, retrieve the object index and use it here.
													 'To see how this should be done, look at the Render Objects demo. Here we won't care about that, and just use "1" since we know the chart is the only object in the file.

		End Sub

		''' <summary>
		''' This event is used when there are linked files, to load them on demand.
		''' </summary>
		Private Sub Workspace_LoadLinkedFile(ByVal sender As Object, ByVal e As LoadLinkedFileEventArgs)
			'In order to reduce the risk of opening any file, in this demo we are going to only open files in the same folder we are working on.
			Dim xls As New XlsFile()
			Dim TemplatePath As String = Path.Combine(Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), ".."), "..")
			xls.Open(Path.Combine(TemplatePath, Path.GetFileName(e.FileName)))

			e.Xls = xls

			'A normal event should end here. Since we need to change the values of the file we loaded in demand, we will do that here.
			xls.SetCellValue(4, 2, GetValue(ChartA1.Text))
			xls.SetCellValue(5, 2, GetValue(ChartA2.Text))
			xls.SetCellValue(6, 2, GetValue(ChartA3.Text))
			xls.SetCellValue(4, 3, GetValue(ChartB1.Text))
			xls.SetCellValue(5, 3, GetValue(ChartB2.Text))
			xls.SetCellValue(6, 3, GetValue(ChartB3.Text))
		End Sub


		'This code is from the "Render objects" demo, and returns the image of a chart.
		Private Function GetChart(ByVal Xls As ExcelFile, ByVal ChartIndex As Integer) As Image
			Dim ChartProps As TShapeProperties = Xls.GetObjectProperties(ChartIndex, True)

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

			Dim dpi As Single = 96 'default screen resolution
			If SizePixels.Height > 0 AndAlso SizePixels.Width > 0 Then
				Dim AspectX As Double = CDbl(chartBox.Width) / SizePixels.Width
				Dim AspectY As Double = CDbl(chartBox.Height) / SizePixels.Height

				Dim Aspect As Double = Math.Max(AspectX, AspectY)
				'Make the dpi adjust the screen resolution and the size of the form.
				dpi = CSng(96 * Aspect)
				If dpi < 20 Then
					dpi = 20
				End If
				If dpi > 500 Then
					dpi = 500
				End If
			End If

			Return Xls.RenderObject(ChartIndex, dpi, ChartProps, SmoothingMode.AntiAlias, InterpolationMode.HighQualityBicubic, True, True, Origin, ImageDimensions, SizePixels)


		End Function

	End Class

End Namespace

