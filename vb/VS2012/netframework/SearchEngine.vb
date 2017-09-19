Imports System.IO
Imports System.Text
Imports System.Globalization
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Reflection

Imports FlexCel.Core
Imports FlexCel.XlsAdapter

Namespace MainDemo
	''' <summary>
	''' A small search engine for finding information in the always-growing number of demos.
	''' This is not production-quality code, just a fast and dirty implementation, but feel free to use
	''' this code in your own projects. 
	''' </summary>
	Public Class SearchEngine
		#Region "Private data"

		Private MainPath As String
		Private FMainException As Exception
		Private FInitialized As Boolean

		Private WordTable As DataTable
		Private WordView As DataView

		Private Const DataFile As String = "keywords.dat"
		Private Const KeywordTable As String = "Keywords"
		Private Const WordColumn As String = "word"
		Private Const ModuleColumn As String = "module"
		Private ReadOnly WordDelim() As Char = { " "c, "."c, "/"c, "\"c }
		#End Region

		#Region "Constructor And Indexing"
		Public Sub New(ByVal aMainPath As String)
			MainPath = aMainPath
			FInitialized = False
		End Sub

		Public Sub Index()
			Try
				Dim ModulesPath As String = Path.GetFullPath(Path.Combine(Path.Combine(Path.Combine(MainPath, ".."), ".."), "Modules"))
				Dim ConfigFile As String = Path.Combine(MainPath, DataFile)

				Try
					WordTable = New DataTable(KeywordTable)
					WordTable.PrimaryKey = New DataColumn() { WordTable.Columns.Add(WordColumn, GetType(String)) }
					WordTable.Columns.Add(ModuleColumn, GetType(ModuleList))

					Dim Loaded As Boolean = False
					If File.Exists(ConfigFile) Then
						Loaded = LoadData(ConfigFile)
					End If

					If Not Loaded Then
						Crawl(ModulesPath)
						SaveData(ConfigFile)
					End If

					WordView = New DataView(WordTable)
				Catch
					File.Delete(ConfigFile)
					Throw
				End Try

				FInitialized = True
			Catch ex As Exception 'this method is designed to run in a thread, so we will not pass exceptions.
				FMainException = ex
			End Try
		End Sub

		Friend ReadOnly Property MainException() As Exception
			Get
				Return FMainException
			End Get
		End Property

		Friend ReadOnly Property Initialized() As Boolean
			Get
				Return FInitialized
			End Get
		End Property

		#End Region

		#Region "Search interface"

		Public Function Search(ByVal words As String) As Dictionary(Of String, String)
			Dim w() As String = words.Split(WordDelim)

			Dim Result As Dictionary(Of String, String) = Nothing

			For Each s As String In w
				Dim s1 As String = s.Trim().ToUpper()
				If s1.Length <= 0 Then
					Continue For
				End If
				s1 = s1.Replace("'", "") 'Avoid escape inside the like expression.
				Dim filter As String = WordColumn & " like '%" & s1 & "%'"
				WordView.RowFilter = filter

				If WordView.Count > 100 Then 'Do not bother filtering by this keyword, too many entries.
					Continue For
				End If


				Dim WordModules As New Dictionary(Of String, String)()
				For i As Integer = 0 To WordView.Count - 1
					Dim dv As DataRowView = WordView(i)
					Dim value As Object = dv(ModuleColumn)

					Dim ht As Dictionary(Of String, String) = CType(value, Dictionary(Of String, String))
					For Each [module] As String In ht.Keys
						WordModules([module]) = [module]
					Next [module]
				Next i

				If Result Is Nothing Then
					Result = WordModules
				Else '"And" words together.
					Dim keys(Result.Keys.Count - 1) As String
					Result.Keys.CopyTo(keys, 0)
					For Each key As String In keys
						If Not WordModules.ContainsKey(key) Then
							Result.Remove(key)
						End If
					Next key
				End If

				If Result.Count = 0 Then 'no need to keep on filtering.
					Return Result
				End If
			Next s
			Return Result

		End Function
		#End Region

		#Region "Implementation"
		Private Sub Crawl(ByVal RelativePath As String)
			Dim Parent As New DirectoryInfo(RelativePath)
			For Each Child As DirectoryInfo In Parent.GetDirectories()
				Crawl(Child.FullName)
			Next Child

			For Each file As FileInfo In Parent.GetFiles("*.rtf")
				AddRtfFile(file.FullName)
			Next file

			For Each file As FileInfo In Parent.GetFiles("*.cs")
				AddTxtFile(file.FullName)
			Next file

			For Each file As FileInfo In Parent.GetFiles("*.vb")
				AddTxtFile(file.FullName)
			Next file

			For Each file As FileInfo In Parent.GetFiles("*.pas")
				AddTxtFile(file.FullName)

			Next file
			For Each file As FileInfo In Parent.GetFiles("*.vb")
				AddTxtFile(file.FullName)
			Next file

			For Each file As FileInfo In Parent.GetFiles("*.txt")
				AddTxtFile(file.FullName)
			Next file

			For Each file As FileInfo In Parent.GetFiles("*.xls")
				AddXlsFile(file.FullName)
			Next file

		End Sub

		Private Sub AddRtfFile(ByVal FileName As String)
			'This implements a *really* basic parser, but it doesn't matter for this use.
			Using sr As New StreamReader(FileName)
				Dim key As Integer
				key = sr.Read()
				Do While key > 0
					If key = AscW("}"c) OrElse key = AscW("{"c) Then
						key = sr.Read()
						Continue Do
					End If
					If key = AscW("\"c) Then
						SkipCommand(sr)
						key = sr.Read()
						Continue Do
					End If

					If Char.IsLetterOrDigit(ChrW(key)) Then
						GetWord(ChrW(key), sr, FileName)
						key = sr.Read()
						Continue Do
					End If
					key = sr.Read()
				Loop
			End Using
		End Sub

		'This method is too naive, it will ignore parameters. This means that in text like:
		' "\fcharset0 Garamond;" Garamond will be considered a word. Again, not a big problem here.
		' What matters more here is speed, and this is faster than using a RichTextBox
		Private Sub SkipCommand(ByVal sr As StreamReader)
			Dim key As Integer
			key = sr.Read()
			Do While key > 0
				If key = AscW(" "c) Then
					Return
				End If
				key = sr.Read()
			Loop
		End Sub

		Private Sub GetWord(ByVal first As Char, ByVal sr As StreamReader, ByVal FileName As String)
			Dim key As Integer
			Dim sb As New StringBuilder()
			sb.Append(first)
			key = sr.Read()
			Do While key > 0
				If Char.IsLetterOrDigit(ChrW(key)) Then
					sb.Append(ChrW(key))
				Else
					AddWord(sb.ToString(), FileName)
					Return
				End If

				key = sr.Read()
			Loop
		End Sub

		Private Sub AddTxtFile(ByVal FileName As String)
			Using sr As New StreamReader(FileName)
				Dim line As String
				line = sr.ReadLine()
				Do While line IsNot Nothing
					Dim words() As String = line.Split(WordDelim)
					For Each word As String In words
						AddWord(word, FileName)
					Next word
					line = sr.ReadLine()
				Loop
			End Using
		End Sub

		Private Sub AddXlsFile(ByVal FileName As String)
			Dim xls As New XlsFile()
			Try
				xls.Open(FileName)
			Catch ex As FlexCelXlsAdapterException
				If ex.ErrorCode = XlsErr.ErrInvalidPassword Then
					Return
				End If
				Throw
			End Try

			For sheet As Integer = 1 To xls.SheetCount
				xls.ActiveSheet = sheet
				For r As Integer = 1 To xls.RowCount
					For cindex As Integer = 1 To xls.ColCountInRow(r)
						Dim XF As Integer = -1
						Dim cell As Object = xls.GetCellValueIndexed(r, cindex, XF)
						AddWord(Convert.ToString(cell), FileName) 'we could use TFlxNumberFormat.FormatValue() here, but we don't care about formatted values for searching.
					Next cindex
				Next r
			Next sheet

		End Sub

		Private Sub AddWord(ByVal word As String, ByVal [module] As String)
			Dim Trimmed As String = word.Trim().ToUpper()
			If Trimmed.Length <= 2 Then 'Filter small words. we need 3, for things like .net or asp, or com.
				Return
			End If

			Dim dr As DataRow = WordTable.Rows.Find(Trimmed)
			If dr Is Nothing Then
				Dim [Mod] As New ModuleList()
				[Mod].Add([module], [module])
				WordTable.Rows.Add(New Object() { Trimmed, [Mod] })
			Else
				CType(dr(ModuleColumn), ModuleList)([module]) = [module]
			End If

		End Sub


		#End Region

		#Region "Save Dataset"

		Private Function FlexCelVersion() As String
			Dim asm As System.Reflection.Assembly = System.Reflection.Assembly.GetAssembly(GetType(XlsFile))
			Return asm.GetName().Version.ToString()
		End Function

		Private Sub SaveData(ByVal filename As String)
			Using fs As New FileStream(filename, FileMode.Create)
				Dim bin As New BinaryFormatter()
				bin.Serialize(fs, FlexCelVersion())
				bin.Serialize(fs, WordTable.Rows.Count)
				For Each dr As DataRow In WordTable.Rows
					bin.Serialize(fs, dr(WordColumn))

					Dim list As ModuleList = CType(dr(ModuleColumn), ModuleList)
					bin.Serialize(fs, list.Count)
					For Each key As String In list.Keys
						bin.Serialize(fs, key)
					Next key

				Next dr
			End Using
		End Sub

		Private Function LoadData(ByVal filename As String) As Boolean
			Try
				Using fs As New FileStream(filename, FileMode.Open)
					If fs.Length <= 0 Then
						Return False
					End If
					Dim bin As New BinaryFormatter()
					Dim Version As String = CStr(bin.Deserialize(fs))
					If Version <> FlexCelVersion() Then 'if this is a new version, regenerate the index.
						Return False
					End If

					Dim Entries As Integer = CInt(Fix(bin.Deserialize(fs)))

					For i As Integer = 0 To Entries - 1
						Dim word As String = CStr(bin.Deserialize(fs))

						Dim modulecount As Integer = CInt(Fix(bin.Deserialize(fs)))
						Dim list As New ModuleList(modulecount)
						For k As Integer = 0 To modulecount - 1
							Dim m As String = CStr(bin.Deserialize(fs))
							list.Add(m, m)
						Next k

						WordTable.Rows.Add(New Object() { word, list })

					Next i

					Debug.Assert(Entries = WordTable.Rows.Count)
				End Using
			Catch
				Return False
			End Try
			Return True
		End Function

		#End Region
	End Class

	Friend Class ModuleList
		Inherits Dictionary(Of String, String)

		Public Sub New()
			MyBase.New(StringComparer.Create(CultureInfo.InvariantCulture, True))
		End Sub

		Public Sub New(ByVal Capacity As Integer)
			MyBase.New(Capacity, StringComparer.Create(CultureInfo.InvariantCulture, True))
		End Sub

	End Class
End Namespace
