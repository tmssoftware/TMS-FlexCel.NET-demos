Imports System.Globalization

Imports FlexCel.Report
Imports System.Collections


Namespace VirtualDatasets
	''' <summary>
	''' This class implements the complete functionality needed to run a FlexCelReport from an array of objects.
	''' Sorting/Filtering on the config sheet is allowed for this datasource, but they are not implemented on an efficient way.
	''' If you do not have an efficient way to do those things (using indexes) and you plan to do them, you will probably get better performance 
	''' using Datasets.
	''' </summary>
	Public Class ComplexVirtualArrayDataSource
		Inherits VirtualDataTable

		#Region "Private variables"
		Private FData()() As Object
		Private FColumnCaptions() As String
		#End Region

		#Region "Constructors"
		Public Sub New(ByVal aCreatedBy As VirtualDataTable, ByVal aData()() As Object, ByVal aColumnCaptions() As String, ByVal aTableName As String)
			MyBase.New(aTableName, aCreatedBy)
			FData = aData
			FColumnCaptions = aColumnCaptions
		End Sub
		#End Region

		#Region "Columns"
		Public Overrides ReadOnly Property ColumnCount() As Integer
			Get
				Return FColumnCaptions.Length
			End Get
		End Property

		Public Overrides Function GetColumn(ByVal columnName As String) As Integer
			'not very optimized method, but this is just a demo.
			If columnName Is Nothing Then
				Return -1
			End If
			columnName = columnName.Trim()
			For i As Integer = 0 To FColumnCaptions.Length - 1
				If String.Compare(FColumnCaptions(i), columnName, True) = 0 Then
					Return i
				End If
			Next i
			Return -1
		End Function

		Public Overrides Function GetColumnName(ByVal columnIndex As Integer) As String
			Return FColumnCaptions(columnIndex)
		End Function

		Public Overrides Function GetColumnCaption(ByVal columnIndex As Integer) As String
			Return FColumnCaptions(columnIndex)
		End Function

		#End Region

		#Region "Settings"
		Public Overrides ReadOnly Property Locale() As System.Globalization.CultureInfo
			Get
				Return CultureInfo.CurrentCulture
			End Get
		End Property
		#End Region

		#Region "Create State"
		Public Overrides Function CreateState(ByVal sort As String, ByVal masterDetailLinks() As TMasterDetailLink, ByVal splitLink As TSplitLink) As VirtualDataTableState
			Return New ComplexVirtualArrayDataSourceState(Me, sort, masterDetailLinks, splitLink)
		End Function

		#End Region

		#Region "Data"
		Public ReadOnly Property Data() As Object()()
			Get
				Return FData
			End Get
		End Property

		Public Overrides Function FilterData(ByVal newDataName As String, ByVal rowFilter As String) As VirtualDataTable
			If rowFilter Is Nothing OrElse rowFilter.Length = 0 Then
				Return New ComplexVirtualArrayDataSource(Me, FData, FColumnCaptions, newDataName) 'no need to copy the data since it is invariant.
			End If

			Dim Relationship As RelationshipType = RelationshipType.Equals
			'on this demo we will only support filters of the type: "field = value", "field > value" or "field < value"
			Dim filteredData() As String = rowFilter.Split("="c)
			If filteredData Is Nothing OrElse filteredData.Length <> 2 Then
				filteredData = rowFilter.Split("<"c)
				If filteredData Is Nothing OrElse filteredData.Length <> 2 Then
					filteredData = rowFilter.Split(">"c)
					If filteredData Is Nothing OrElse filteredData.Length <> 2 Then
						Throw New Exception("Filter """ & rowFilter & """ is invalid. The dataset """ & TableName & """ only supports filters of the type ""field =/>/< value"".")
					Else
						Relationship = RelationshipType.BiggerThan
					End If
				Else
					Relationship = RelationshipType.LessThan
				End If

			End If

			Dim ColIndex As Integer = GetColumn(filteredData(0).Trim())
			If ColIndex < 0 Then
				Throw New Exception("Filter """ & rowFilter & """ is invalid. Can not find column """ & filteredData(0).Trim() & """")
			End If

			'Remember, this is only a demo to show what to do in this event. This code is not good code to use on 
			'a real application!. You should use some index here to find the data or this would crawl on large objects.

			Dim Result As New List(Of Object())()
			Dim SearchValue As String = filteredData(1).Trim()
			Dim Searchfloat As Double = 0
			If Relationship <> RelationshipType.Equals Then 'when relationship is not equals, we will assume columns are numbers. This is because we do not have any type definition for our columns on this simple example.
				Searchfloat = Convert.ToDouble(filteredData(1).Trim())
			End If

			For Each Row As Object() In Data
				Select Case Relationship
					Case RelationshipType.Equals
						If Convert.ToString(Row(ColIndex)) = SearchValue Then
							Result.Add(Row) 'remember, data is invariant, so we do not need to clone Row.
						End If
					Case RelationshipType.LessThan
						If Convert.ToDouble(Row(ColIndex)) < Searchfloat Then
							Result.Add(Row) 'remember, data is invariant, so we do not need to clone Row.
						End If
					Case RelationshipType.BiggerThan
						If Convert.ToDouble(Row(ColIndex)) > Searchfloat Then
							Result.Add(Row) 'remember, data is invariant, so we do not need to clone Row.
						End If
				End Select
			Next Row

			Return New ComplexVirtualArrayDataSource(Me, Result.ToArray(), FColumnCaptions, newDataName)

		End Function

		Public Overrides Function GetDistinct(ByVal newDataName As String, ByVal filterFields() As Integer) As VirtualDataTable
			If filterFields Is Nothing OrElse filterFields.Length = 0 Then
				Return New ComplexVirtualArrayDataSource(Me, FData, FColumnCaptions, newDataName) 'no need to copy the data since it is invariant.
			End If

			Dim Result As New Dictionary(Of Object() , Object())()
			Dim Keys(filterFields.Length - 1) As Object
			For Each Row As Object() In FData
				For i As Integer = 0 To filterFields.Length - 1
					Keys(i) = Row(filterFields(i))
				Next i
				Result(Keys) = Keys
			Next Row

			Dim R(Data.Length - 1)() As Object
			Result.Keys.CopyTo(R, 0)

			Dim NewColumnCaptions(filterFields.Length - 1) As String
			For i As Integer = 0 To filterFields.Length - 1
				NewColumnCaptions(i) = FColumnCaptions(filterFields(i))
			Next i
			Return New ComplexVirtualArrayDataSource(Me, R, NewColumnCaptions, newDataName)

		End Function

		#End Region

	End Class

	Public Class ComplexVirtualArrayDataSourceState
		Inherits VirtualDataTableState

		#Region "Privates"
		Private SortedData()() As Object
		Private FilteredData As List(Of Object())
		#End Region

		#Region "Constructors"
		Public Sub New(ByVal aTableData As ComplexVirtualArrayDataSource, ByVal sort As String, ByVal masterDetailLinks() As TMasterDetailLink, ByVal splitLink As TSplitLink)
			MyBase.New(aTableData)
			If sort Is Nothing OrElse sort.Trim().Length = 0 Then
				SortedData = aTableData.Data 'no need to clone, this is invariant.
			Else
				SortedData = CType(aTableData.Data.Clone(), Object()())
				Dim sortcolumn As Integer = aTableData.GetColumn(sort)
				If sortcolumn < 0 Then
					Throw New Exception("Can not find column """ & sort & """ in dataset """ & TableName)
				End If
				Array.Sort(SortedData, New ArrayComparer(sortcolumn))
			End If


			'here we should use the data in masterdetaillinks and splitlink to create indexes to make the FilteredrowCount and MoveMasterRecord methods faster.
			'on this demo we are not going to do it.
			If (masterDetailLinks IsNot Nothing AndAlso masterDetailLinks.Length > 0) OrElse splitLink IsNot Nothing Then
				FilteredData = New List(Of Object())()
			End If
		End Sub
		#End Region

		#Region "Data"
		Private ReadOnly Property Data() As Object()()
			Get
				Return CType(TableData, SimpleVirtualArrayDataSource).Data
			End Get
		End Property

		''' <summary>
		''' Remember that this method should be fast!
		''' </summary>
		Public Overrides ReadOnly Property RowCount() As Integer
			Get
				If FilteredData Is Nothing Then
					Return SortedData.Length
				End If
				Return FilteredData.Count
			End Get
		End Property

		Public Overrides Function GetValue(ByVal column As Integer) As Object
			If FilteredData Is Nothing Then
				Return SortedData(Position)(column)
			End If
			Return CType(FilteredData(Position), Object())(column)
		End Function

		Public Overrides Function GetValue(ByVal row As Integer, ByVal column As Integer) As Object
			If FilteredData Is Nothing Then
				Return SortedData(row)(column)
			End If
			Return CType(FilteredData(row), Object())(column)
		End Function
		#End Region

		#Region "Move"
		Public Overrides Sub MoveFirst()
			'No need to do anything in an array, since "Position" is moved for us. 
			'If we had an IEnumerator for example as data backend, we would reset it here.
		End Sub

		Public Overrides Sub MoveNext()
			'No need to do anything in an array, since "Position" is moved for us. 
			'If we had an IEnumerator for example as data backend, we would move it here.
		End Sub

		#End Region


		#Region "Relationships"
		Public Overrides Sub MoveMasterRecord(ByVal masterDetailLinks() As TMasterDetailLink, ByVal splitLink As TSplitLink)
			'Here we need to modify FilteredData to contain only those records on FilteredData that are visible on this state.
			'If we created indexes on this method constructor's, this can be done fast.

			'on this example, as always, we will use a very slow method. The idea of this demo is not to demostrate how to do efficient code
			' (you probably have efficient methods on the bussines objects you are wrapping), but how to override the methods.

			If FilteredData Is Nothing Then 'this dataset is not on master-detail relationship and does not have any split relationship either.
				Return
			End If


			FilteredData.Clear()
			Dim ChildColumn(masterDetailLinks.Length - 1) As Integer
			For i As Integer = 0 To masterDetailLinks.Length - 1
				ChildColumn(i) = TableData.GetColumn(masterDetailLinks(i).ChildFieldName)
			Next i

			Dim SplitPos As Integer = 0
			Dim StartRow As Integer = 0
			If splitLink IsNot Nothing Then
				StartRow = splitLink.SplitCount * splitLink.ParentDataSource.Position
			End If

			For r As Integer = 0 To SortedData.Length - 1
				Dim Row() As Object = SortedData(r)
				Dim RowApplies As Boolean = True
				For i As Integer = 0 To masterDetailLinks.Length - 1
					Dim key As Object = masterDetailLinks(i).ParentDataSource.GetValue(masterDetailLinks(i).ParentField)
					If Convert.ToString(Row(ChildColumn(i))) <> Convert.ToString(key) Then
						RowApplies = False
						Exit For
					End If
				Next i

				If Not RowApplies Then 'the row does not fit this master detail relationship.
					Continue For
				End If

				SplitPos += 1
				If SplitPos <= StartRow Then 'we are not filling the correct split slot.
					Continue For
				End If

				FilteredData.Add(Row)
				If splitLink IsNot Nothing AndAlso FilteredData.Count >= splitLink.SplitCount Then
					Return
				End If
			Next r


		End Sub

		Public Overrides Function FilteredRowCount(ByVal masterDetailLinks() As TMasterDetailLink) As Integer
			Dim Result As Integer = 0

			Dim ChildColumn(masterDetailLinks.Length - 1) As Integer
			For i As Integer = 0 To masterDetailLinks.Length - 1
				ChildColumn(i) = TableData.GetColumn(masterDetailLinks(i).ChildFieldName)
			Next i

			For r As Integer = 0 To SortedData.Length - 1
				Dim Row() As Object = SortedData(r)
				Dim RowApplies As Boolean = True
				For i As Integer = 0 To masterDetailLinks.Length - 1
					Dim key As Object = masterDetailLinks(i).ParentDataSource.GetValue(masterDetailLinks(i).ParentField)
					If Convert.ToString(Row(ChildColumn(i))) <> Convert.ToString(key) Then
						RowApplies = False
						Exit For
					End If
				Next i

				If Not RowApplies Then 'the row does not fit this master detail relationship.
					Continue For
				End If

				Result += 1
			Next r

			Return Result

		End Function


		#End Region

	End Class

	Friend Enum RelationshipType
		Equals
		LessThan
		BiggerThan
	End Enum

	Public Class ArrayComparer
		Implements IComparer(Of Object())

		Private Column As Integer
		Private Shared Comparer As New CaseInsensitiveComparer()

		Public Sub New(ByVal aColumn As Integer)
			Column = aColumn
		End Sub

		Private Function IComparerGeneric_Compare(ByVal x() As Object, ByVal y() As Object) As Integer Implements IComparer(Of Object()).Compare
			Return (Comparer.Compare(x(Column), y(Column)))
		End Function

	End Class


End Namespace
