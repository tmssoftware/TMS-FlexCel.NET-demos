Imports System.Globalization

Imports FlexCel.Report

Namespace VirtualDatasets
	''' <summary>
	''' This class implements the minimum functionality needed to run a FlexCelReport from an array of objects.
	''' No Sorting/Filtering on the config sheet is allowed for this datasource, and you can not use it on master detail relationships.
	''' </summary>
	Public Class SimpleVirtualArrayDataSource
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
			Return New SimpleVirtualArrayDataSourceState(Me)
		End Function

		#End Region

		#Region "Data"
		Public ReadOnly Property Data() As Object()()
			Get
				Return FData
			End Get
		End Property
		#End Region

		#Region "Filter"
		''' <summary>
		''' Even when we don't implement filter, we need to return a new instance when rowFilter is null.
		''' </summary>
		''' <param name="newDataName"></param>
		''' <param name="rowFilter"></param>
		''' <returns></returns>
		Public Overrides Function FilterData(ByVal newDataName As String, ByVal rowFilter As String) As VirtualDataTable
			If String.IsNullOrEmpty(rowFilter) Then
				Return New SimpleVirtualArrayDataSource(Me, Data, FColumnCaptions, TableName)
			End If
			Return MyBase.FilterData(newDataName, rowFilter)
		End Function
		#End Region

	End Class

	Public Class SimpleVirtualArrayDataSourceState
		Inherits VirtualDataTableState

		#Region "Constructors"
		Public Sub New(ByVal aTableData As SimpleVirtualArrayDataSource)
			MyBase.New(aTableData)
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
				Return Data.Length
			End Get
		End Property

		Public Overrides Function GetValue(ByVal column As Integer) As Object
			Return Data(Position)(column)
		End Function

		Public Overrides Function GetValue(ByVal row As Integer, ByVal column As Integer) As Object
			Return Data(row)(column)
		End Function

		Public Overrides Sub MoveFirst()
			'No need to do anything in an array, since "Position" is moved for us. 
			'If we had an IEnumerator for example as data backend, we would reset it here.
		End Sub

		Public Overrides Sub MoveNext()
			'No need to do anything in an array, since "Position" is moved for us. 
			'If we had an IEnumerator for example as data backend, we would move it here.
		End Sub

		#End Region

	End Class

End Namespace
