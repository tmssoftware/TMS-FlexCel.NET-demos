Imports System.Collections
Imports System.ComponentModel
Imports System.IO

Imports FlexCel.Core
Imports FlexCel.XlsAdapter

Namespace ObjectExplorer
	''' <summary>
	''' Object explorer.
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
		#End Region

		Private Sub btnExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.Click
			Close()
		End Sub

		Private Sub btnOpenFile_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOpenFile.Click
			If openFileDialog1.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
				Return
			End If

			Xls = New XlsFile()
			Xls.Open(openFileDialog1.FileName)


			cbSheet.Items.Clear()
			Dim ActSheet As Integer = Xls.ActiveSheet
			For i As Integer = 1 To Xls.SheetCount
				Xls.ActiveSheet = i
				cbSheet.Items.Add(Xls.SheetName)
			Next i
			Xls.ActiveSheet = ActSheet
			cbSheet.SelectedIndex = ActSheet - 1

			FillListBox()
		End Sub

		Private Sub FillListBox()
			lblObjects.Text = openFileDialog1.FileName
			dataGrid.DataSource = Nothing

			ObjTree.BeginUpdate()
			Try
				ObjTree.Nodes.Clear()

				For i As Integer = 1 To Xls.ObjectCount
					Dim ShapeProps As TShapeProperties = Xls.GetObjectProperties(i, True)
					Dim s As String = "Object " & i.ToString()
					If ShapeProps.ShapeName IsNot Nothing Then
						s = ShapeProps.ShapeName
					End If

					Dim RootNode As New TreeNode(s)
					FillNodes(ShapeProps, RootNode)


					ObjTree.Nodes.Add(RootNode)
				Next i
			Finally
				ObjTree.EndUpdate()
			End Try
		End Sub

		Private Sub FillNodes(ByVal ShapeProps As TShapeProperties, ByVal Node As TreeNode)
			Node.Tag = ShapeProps 'In this simple demo we will use the tag property to store the Shape properties. This is not indented for 'real' use.


			For i As Integer = 1 To ShapeProps.ChildrenCount
				Dim ChildProps As TShapeProperties = ShapeProps.Children(i)
				Dim ShapeName As String = ChildProps.ShapeName
				If ShapeName Is Nothing Then
					ShapeName = "Object " & i.ToString()
				End If
				Dim Child As New TreeNode(ShapeName)
				FillNodes(ChildProps, Child)
				Node.Nodes.Add(Child)
			Next i
		End Sub

		Private Sub btnOpen_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnShowInExcel.Click
			If Xls Is Nothing Then
				MessageBox.Show("There is no open file")
				Return
			End If
			System.Diagnostics.Process.Start(Xls.ActiveFileName)
		End Sub


		Private Sub btnSaveImage_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSaveAsImage.Click
			If PreviewBox.Image Is Nothing Then
				MessageBox.Show("There is no selected image to save", "Error")
				Return
			End If
			If saveImageDialog.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
				Return
			End If

			PreviewBox.Image.Save(saveImageDialog.FileName)
		End Sub

		Private Sub btnInfo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnInfo.Click
			MessageBox.Show("Object Explorer allows to explore inside the objects in an Excel file." & vbLf & "Objects in xls files are hierachily distributed, you can have two objects grouped as a third object, " & "and this hierarchy is shown in the 'Objects' pane at the left. The properties for the selected object are displayed at the 'Object properties' pane.")
		End Sub

		Private Sub ObjTree_AfterSelect(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles ObjTree.AfterSelect
			RenderNode(e.Node)
			FillProperties(e.Node)
		End Sub


		Private Sub RenderNode(ByVal Node As TreeNode)
			If Xls Is Nothing Then
				PreviewBox.Image = Nothing
				Return
			End If

			Dim t As TreeNode = Node
			If t Is Nothing Then
				PreviewBox.Image = Nothing
				Return
			End If

			Do While t.Parent IsNot Nothing 'Only root level objects will be rendered.
				t = t.Parent
			Loop

			If t.Index + 1 > Xls.ObjectCount Then
				PreviewBox.Image = Nothing
				Return
			End If

			PreviewBox.Image = Xls.RenderObject(t.Index + 1)


		End Sub

		Private Sub FillProperties(ByVal Node As TreeNode)
			lblObjName.Text = "Name:"
			lblObjText.Text = "Text:"
			Dim Props As TShapeProperties = CType(Node.Tag, TShapeProperties)
			If Props Is Nothing Then
				dataGrid.DataSource = Nothing
				Return
			End If

			Dim ShapeOptions As TShapeOptionList = (TryCast(Node.Tag, TShapeProperties)).ShapeOptions
			If ShapeOptions Is Nothing Then
				dataGrid.DataSource = Nothing
				Return
			End If

			lblObjName.Text = "Name: " & Props.ShapeName
			lblObjText.Text = "Text: " & Convert.ToString(Props.Text)

			Dim ShapeOpts As New ArrayList()
			For Each opt As KeyValuePair(Of TShapeOption, Object) In ShapeOptions
				ShapeOpts.Add(New KeyValue(opt.Key, ShapeOptions))
			Next opt
			dataGrid.DataSource = ShapeOpts
		End Sub

		Private Sub cbSheet_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbSheet.SelectedIndexChanged
			Xls.ActiveSheet = cbSheet.SelectedIndex + 1
			FillListBox()
		End Sub

		Private Sub btnStretchPreview_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnStretchPreview.Click
			If btnStretchPreview.Checked Then
				PreviewBox.SizeMode = PictureBoxSizeMode.StretchImage
			Else
				PreviewBox.SizeMode = PictureBoxSizeMode.Normal
			End If

		End Sub

	End Class

	Friend Class KeyValue
		Private FKey As String
		Private FAs1616 As String
		Private FAsLong As String
		Private FAsString As String

		Public Sub New(ByVal aKey As TShapeOption, ByVal List As TShapeOptionList)
			FKey = Convert.ToString(aKey)
			FAs1616 = Convert.ToString(List.As1616(aKey, 0))
			FAsLong = Convert.ToString(List.AsLong(aKey, 0))
			FAsString = List.AsUnicodeString(aKey, "")
		End Sub

		Public Property Key() As String
			Get
				Return FKey
			End Get
			Set(ByVal value As String)
				FKey = value
			End Set
		End Property
		Public Property As1616() As String
			Get
				Return FAs1616
			End Get
			Set(ByVal value As String)
				FAs1616 = value
			End Set
		End Property
		Public Property AsLong() As String
			Get
				Return FAsLong
			End Get
			Set(ByVal value As String)
				FAsLong = value
			End Set
		End Property
		Public Property AsString() As String
			Get
				Return FAsString
			End Get
			Set(ByVal value As String)
				FAsString = value
			End Set
		End Property

		Public Overrides Function ToString() As String
			Return Key
		End Function

	End Class
End Namespace
