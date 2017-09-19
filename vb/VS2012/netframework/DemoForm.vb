Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports System.Globalization
Imports System.Resources
Imports System.Threading

Namespace MainDemo
	''' <summary>
	''' Demo Browser for FlexCel. This application will run all the other demos available.
	''' </summary>
	Partial Public Class DemoForm
		Inherits System.Windows.Forms.Form

		Public Sub New()
			InitializeComponent()
			ResizeToolbar(mainToolbar)
			descriptionText.BackColor = System.Drawing.SystemColors.Window 'It is lost when converting to .NET 2.0
			CleanSearchbox()
			LoadModules()
			FilterTree(Nothing)
		End Sub

		Private Sub ResizeToolbar(ByVal toolbar As ToolStrip)

			Using gr As Graphics = CreateGraphics()
				Dim xFactor As Double = gr.DpiX / 96.0
				Dim yFactor As Double = gr.DpiY / 96.0
				toolbar.ImageScalingSize = New Size(CInt(Fix(24 * xFactor)), CInt(Fix(24 * yFactor)))
				toolbar.Width = 0 'force a recalc of the buttons.
			End Using

		End Sub


		#Region "Global constants."
		Private PathToExe As String = Path.Combine("bin", "Debug")
		Private ExtLaunch As String = ".xls"
		Private ExtTemplate As String = ".template.xls"
		Private ExtCsProject As String = ".csproj"
		Private ExtVbProject As String = ".vbproj"
		Private ExtPrismProject As String = ".oxygene"
		Private ExtLocation As String = ".location.txt"

		Private Finder As SearchEngine
		Private MainNode As TTreeNode
		#End Region

		Private Sub LoadModules()
			Dim MainPath As String = Path.GetFullPath(Path.Combine(Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), ".."), ".."))
			MainNode = New TTreeNode(Text, Path.Combine(MainPath, "MainDemo"))

			Dim subdirectoryEntries() As String = Directory.GetDirectories(Path.Combine(MainPath, "Modules"))
			For Each subdirectory As String In subdirectoryEntries
				LoadModule(Path.Combine(MainPath, "Modules"), subdirectory, MainNode)
			Next subdirectory
		End Sub

		Private Sub LoadModule(ByVal MainPath As String, ByVal modulePath As String, ByVal node As TTreeNode)
			Dim LinkFile As String = Path.Combine(modulePath, "link.txt")
			If File.Exists(LinkFile) Then
				Using sr As New StreamReader(LinkFile)
					Dim RelPath As String = sr.ReadLine().Replace("\"c, Path.DirectorySeparatorChar)
					modulePath = Path.GetFullPath(Path.Combine(MainPath, RelPath))
				End Using
			End If

			Dim moduleName As String = Path.GetFileName(modulePath)
			Dim shortModule As String = moduleName.Substring(moduleName.IndexOf(".") + 1)
			If moduleName.Length < 1 OrElse moduleName.Chars(0) = "."c Then 'Do not process hidden folders.
				Return
			End If
			If moduleName.IndexOf("."c) < 1 Then 'Do not process folders without the convention xx.name
				Return
			End If

			Dim NodePath As String = Nothing
			If File.Exists(Path.Combine(modulePath, shortModule & ".rtf")) Then
				NodePath = Path.Combine(modulePath, shortModule)
			End If

			Dim NewNode As New TTreeNode(shortModule, NodePath)
			node.Children.Add(NewNode)


			Dim subdirectoryEntries() As String = Directory.GetDirectories(modulePath)
			For Each subdirectory As String In subdirectoryEntries
				LoadModule(MainPath, subdirectory, NewNode)
			Next subdirectory
		End Sub

		Private Sub modulesList_AfterSelect(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles modulesList.AfterSelect
			If e.Node.Tag Is Nothing Then
				descriptionText.Clear()
			Else
				descriptionText.LoadFile((CStr(e.Node.Tag)) & ".rtf")
			End If

			statusBar1.Text = e.Node.FullPath

			btnRunSelected.Enabled = (HasModuleForm()) OrElse (HasFileToLaunch(ExtLaunch) IsNot Nothing) OrElse (HasFileToLaunch(ExtCsProject) IsNot Nothing) OrElse (HasFileToLaunch(ExtVbProject) IsNot Nothing)

			menuRunSelected.Enabled = btnRunSelected.Enabled

			btnViewTemplate.Enabled = HasFileToLaunch(ExtTemplate) IsNot Nothing
			menuViewTemplate.Enabled = btnViewTemplate.Enabled

			btnOpenProject.Enabled = HasFileToLaunch(ExtCsProject) IsNot Nothing OrElse HasFileToLaunch(ExtVbProject) IsNot Nothing OrElse HasFileToLaunch(ExtPrismProject) IsNot Nothing
			menuOpenProject.Enabled = btnOpenProject.Enabled
		End Sub

		Private Sub Exit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles menuExit.Click, btnExit.Click
			Close()
		End Sub

		Private Function HasModuleForm() As Boolean
			Dim Frm As Form = GetModuleForm()
			If Frm Is Nothing Then
				Return False
			End If
			Frm.Dispose()
			Return True
		End Function

		Private Function GetModuleForm() As Form
			Dim mName As String = Nothing
			Dim moduleName As String = GetModuleName(mName)
			If moduleName Is Nothing OrElse (Not File.Exists(moduleName)) Then
				Return Nothing
			End If
			Dim [assembly] As System.Reflection.Assembly = System.Reflection.Assembly.LoadFrom(moduleName)
			Return CType([assembly].CreateInstance(mName & ".mainForm"), Form)
		End Function

		Private Function GetModuleName(ByRef mName As String) As String
			mName = Nothing
			If modulesList.SelectedNode Is Nothing OrElse modulesList.SelectedNode.Tag Is Nothing Then
				Return Nothing
			End If
			Dim mPath As String = Path.Combine(Path.GetDirectoryName((CStr(modulesList.SelectedNode.Tag))), PathToExe)
			mName = Path.GetFileName(CStr(modulesList.SelectedNode.Tag))
			mName = mName.Replace(" ", "")
			Return Path.GetFullPath(Path.Combine(mPath, mName & ".exe"))
		End Function

		Private Function HasFileToLaunch(ByVal extension As String) As String
			If modulesList.SelectedNode Is Nothing OrElse modulesList.SelectedNode.Tag Is Nothing Then
				Return Nothing
			End If
			Dim mPath As String = Path.GetDirectoryName((CStr(modulesList.SelectedNode.Tag)))
			Dim mName As String = Path.GetFileName(CStr(modulesList.SelectedNode.Tag))

			If File.Exists(Path.Combine(mPath, extension.Substring(1) & ExtLocation)) Then
				Using sr As New StreamReader(Path.Combine(mPath, extension.Substring(1) & ExtLocation))
					Return mPath & sr.ReadLine()
				End Using
			End If
			If File.Exists(Path.Combine(mPath, mName & extension)) Then
				Return Path.Combine(mPath, mName & extension)
			End If
			Return Nothing
		End Function

		Private Function IgnoreInMainDemo() As Boolean
			Return IgnoreInMainDemoMessage() IsNot Nothing
		End Function

		Private Function IgnoreInMainDemoMessage() As String
			Dim IgnoreFile As String = HasFileToLaunch(".IgnoreInMainDemo.txt")
			If String.IsNullOrEmpty(IgnoreFile) Then
				Return Nothing
			End If
			Return File.ReadAllText(IgnoreFile)
		End Function


		Private Sub RunSelected_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles menuRunSelected.Click, btnRunSelected.Click
			If IgnoreInMainDemo() Then
				MessageBox.Show(IgnoreInMainDemoMessage())
				Return
			End If

			TryToCompileProject()
			Dim frm As Form = GetModuleForm()
			Try
				If frm Is Nothing Then
					Dim f As String = HasFileToLaunch(ExtLaunch)
					If f IsNot Nothing Then
						System.Diagnostics.Process.Start(f)
					End If
					Return
				End If
				Dim tfrm As Type = frm.GetType()
				Dim autorun As MethodInfo = tfrm.GetMethod("AutoRun")
				If autorun IsNot Nothing Then
					autorun.Invoke(frm, New Object(){})
					Return
				End If

				frm.StartPosition = FormStartPosition.CenterParent
				frm.ShowInTaskbar = False
				frm.ShowDialog()
			Finally
				If frm IsNot Nothing Then
					frm.Dispose()
				End If
			End Try
		End Sub

		Private Sub TryToCompileProject()
			Dim mName As String = Nothing
			Dim moduleName As String = GetModuleName(mName)
			If moduleName IsNot Nothing AndAlso File.Exists(moduleName) Then
				Return
			End If


			Dim CsProj As String = HasFileToLaunch(ExtCsProject)
			If CsProj IsNot Nothing Then
				Builder.Build(CsProj)
			End If

			Dim VbProj As String = HasFileToLaunch(ExtVbProject)
			If VbProj IsNot Nothing Then
				Builder.Build(VbProj)
			End If

		End Sub

		Private Sub ViewTemplate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles menuViewTemplate.Click, btnViewTemplate.Click
			Dim f As String = HasFileToLaunch(ExtTemplate)
			If f IsNot Nothing Then
				System.Diagnostics.Process.Start(f)
			End If

		End Sub

		Private Sub About_Click(ByVal sender As Object, ByVal e As EventArgs) Handles menuAbout.Click, btnAbout.Click
			Using af As New AboutForm()
				af.ShowDialog()
			End Using
		End Sub

		Private Sub descriptionText_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkClickedEventArgs) Handles descriptionText.LinkClicked
			Try
				System.Diagnostics.Process.Start(e.LinkText)
			Catch ex As Exception
				MessageBox.Show(ex.Message)
			End Try
		End Sub

		Private Sub btnOpenProject_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles menuOpenProject.Click, btnOpenProject.Click
			Dim f As String = HasFileToLaunch(ExtCsProject)
			If f IsNot Nothing Then
				System.Diagnostics.Process.Start(f)
				Return
			End If

			f = HasFileToLaunch(ExtVbProject)
			If f IsNot Nothing Then
				System.Diagnostics.Process.Start(f)
				Return
			End If

			f = HasFileToLaunch(ExtPrismProject)
			If f IsNot Nothing Then
				System.Diagnostics.Process.Start(f)
				Return
			End If

		End Sub

		Private Sub btnOpenFolder_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnOpenFolder.Click
			If modulesList.SelectedNode Is Nothing OrElse modulesList.SelectedNode.Tag Is Nothing Then
				Return
			End If
			Dim f As String = Path.GetDirectoryName(CStr(modulesList.SelectedNode.Tag))
			System.Diagnostics.Process.Start(f)

		End Sub

		Private Sub sdSearch_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles sdSearch.TextChanged
			If sdSearch.Tag IsNot Nothing Then
				Return
			End If

			If Finder Is Nothing OrElse (Not Finder.Initialized) Then
				Finder = New SearchEngine(Path.GetDirectoryName(Application.ExecutablePath))
				Dim SearchThread As New Thread(New ThreadStart(AddressOf Finder.Index))
				SearchThread.Start()

				Using Pg As New ProgressDialog()
					Pg.ShowProgress(SearchThread)
					If Finder IsNot Nothing AndAlso Finder.MainException IsNot Nothing Then
						Dim ex As Exception = Finder.MainException
						Finder = Nothing
						Throw ex
					End If
				End Using
			End If

			If String.Compare(sdSearch.Text, "why?", True, CultureInfo.InvariantCulture) = 0 Then
				Answer()
			End If

			Dim FoundModules As Dictionary(Of String, String) = Finder.Search(sdSearch.Text)
			FilterTree(FoundModules)
		End Sub

		Private Sub FilterTree(ByVal FoundModules As Dictionary(Of String, String))
			modulesList.BeginUpdate()
			Try
				Dim OldSelected As TreeNode = modulesList.SelectedNode
				Dim OldSelectedPath As String = Nothing
				If OldSelected IsNot Nothing Then
					OldSelectedPath = Convert.ToString(OldSelected.Tag)
				End If

				modulesList.Nodes.Clear()
				Dim MainTreeNode As New TreeNode(MainNode.NodeName)
				MainTreeNode.Tag = MainNode.NodePath
				modulesList.Nodes.Add(MainTreeNode)
				Dim NewSelected As TreeNode = Nothing
				FilterTree(FoundModules, MainNode, MainTreeNode, OldSelectedPath, NewSelected)
				modulesList.ExpandAll()
				If NewSelected Is Nothing Then
					NewSelected = MainTreeNode
				End If
				modulesList.SelectedNode = NewSelected
				NewSelected.EnsureVisible()
			Finally
				modulesList.EndUpdate()
			End Try
		End Sub

		Private Sub FilterTree(ByVal FoundModules As Dictionary(Of String, String), ByVal ParentNode As TTreeNode, ByVal ParentTreeNode As TreeNode, ByVal OldSelectedPath As String, ByRef NewSelected As TreeNode)
			For Each node As TTreeNode In ParentNode.Children
				If FoundModules Is Nothing OrElse HasKey(FoundModules, Path.GetDirectoryName(node.NodePath)) Then
					Dim NewNode As New TreeNode(node.NodeName)
					NewNode.Tag = node.NodePath
					ParentTreeNode.Nodes.Add(NewNode)
					FilterTree(FoundModules, node, NewNode, OldSelectedPath, NewSelected)
					If node.NodePath = OldSelectedPath Then
						NewSelected = NewNode
					End If
				End If

			Next node
		End Sub

		Private Function HasKey(ByVal FoundModules As Dictionary(Of String, String), ByVal pattern As String) As Boolean
			If pattern Is Nothing Then
				Return False
			End If
			For Each s As String In FoundModules.Keys
				If s.StartsWith(pattern) Then
					Return True
				End If
			Next s
			Return False
		End Function


		Private Sub Answer()
			Dim Answers() As String = { "It was not my fault. I was just following your orders.", "Because that's the way life is. Better go getting used to it.", "The answer is 42. Sometimes.", "If I told you then I would have to kill you.", "It is the user's fault", "I can only answer you after my NDA expires.", "Whatever it is, don't worry. Tomorrow we will look at it and we will laugh.", "Please give me some time to think about it.", "I could tell you, but then where would be the fun?" }

			Dim rnd As New Random()
			MessageBox.Show(Answers(rnd.Next(Answers.Length)))
		End Sub

		Private TxtTypeToSearch As String = "Type to search..." 'this isn't a nice way to show a hint, but it will work for this simple demo, without using a third party control.

		Private Sub sdSearch_Enter(ByVal sender As Object, ByVal e As EventArgs) Handles sdSearch.Enter
			sdSearch.ForeColor = Color.Black
			If sdSearch.Tag IsNot Nothing Then
				sdSearch.Text = ""
				sdSearch.Tag = Nothing
			End If
		End Sub

		Private Sub sdSearch_Leave(ByVal sender As Object, ByVal e As EventArgs) Handles sdSearch.Leave
			CleanSearchbox()
		End Sub

		Private Sub CleanSearchbox()
			sdSearch.ForeColor = Color.Gray
			If String.IsNullOrEmpty(sdSearch.Text) Then
				sdSearch.Tag = "e"
				sdSearch.Text = TxtTypeToSearch
			Else
				sdSearch.Tag = Nothing
			End If
		End Sub

	End Class


	Friend Class TTreeNode
		Public NodeName As String
		Public NodePath As String
		Public Children As List(Of TTreeNode)

		Public Sub New(ByVal aNodeName As String, ByVal aNodePath As String)
			NodeName = aNodeName
			NodePath = aNodePath
			Children = New List(Of TTreeNode)()
		End Sub
	End Class
End Namespace

