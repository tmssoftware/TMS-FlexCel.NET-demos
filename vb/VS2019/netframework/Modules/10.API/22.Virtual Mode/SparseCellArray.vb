Imports System.Text

Namespace VirtualMode
	'''	<summary>
	'''	  This is a simple class that holds cell values. Items are supposed to
	'''	  be entered in sorted order, and it isn't really production-ready, just
	'''	  to be used in a demo.
	'''	</summary>
	Friend Class SparseCellArray
		Private Data As List(Of SparseRow)
		Private FColCount As Integer

		Public Sub New()
			FColCount = 0
		End Sub
		Public Sub AddValue(ByVal Row As Integer, ByVal Col As Integer, ByVal Value As String)
			If Col > FColCount Then
				FColCount = Col
			End If
			If Data Is Nothing Then
				Data = New List(Of SparseRow)()
			End If
			Dim SpRow As New SparseRow(Row)
			Dim Idx As Integer = Data.BinarySearch(SpRow)
			If Idx < 0 Then
				SpRow.CreateData()
				Data.Insert((Not Idx), SpRow)
			Else
				SpRow = Data(Idx)
			End If

			Dim SpCell As New SparseCell(Col, Value)
			Idx = SpRow.Data.BinarySearch(SpCell)
			If Idx < 0 Then
				SpRow.Data.Insert((Not Idx), SpCell)
			Else
				SpRow.Data(Idx) = SpCell
			End If

		End Sub

		Public Function GetValue(ByVal Row As Integer, ByVal Col As Integer) As String
			If Data Is Nothing Then
				Return Nothing
			End If

			Dim SpRow As New SparseRow(Row)
			Dim Idx As Integer = Data.BinarySearch(SpRow)
			If Idx < 0 Then
				Return Nothing
			End If
			SpRow = Data(Idx)
			Dim SpCell As New SparseCell(Col, Nothing)
			Idx = SpRow.Data.BinarySearch(SpCell)
			If Idx < 0 Then
				Return Nothing
			End If
			Return SpRow.Data(Idx).Value
		End Function

		Public ReadOnly Property ColCount() As Integer
			Get
				Return FColCount
			End Get
		End Property
		Public ReadOnly Property RowCount() As Integer
			Get
				If Data Is Nothing OrElse Data.Count = 0 Then
					Return 0
				End If
				Return Data(Data.Count - 1).Row
			End Get
		End Property
	End Class

	Friend Structure SparseRow
		Implements IComparable(Of SparseRow)

		Public Row As Integer
		Public Data As List(Of SparseCell)

		Public Sub New(ByVal aRow As Integer)
			Row = aRow
			Data = Nothing
		End Sub

		Public Sub CreateData()
			Data = New List(Of SparseCell)()
		End Sub

		#Region "IComparable<SparseRow> Members"

		Public Function CompareTo(ByVal other As SparseRow) As Integer Implements IComparable(Of SparseRow).CompareTo
			Return Row.CompareTo(other.Row)
		End Function

		#End Region
	End Structure

	Friend Class SparseCell
		Implements IComparable(Of SparseCell)

		Public Col As Integer
		Public Value As String

		Public Sub New(ByVal aCol As Integer, ByVal aValue As String)
			Col = aCol
			Value = aValue
		End Sub

		#Region "IComparable<SparseCell> Members"

		Public Function CompareTo(ByVal other As SparseCell) As Integer Implements IComparable(Of SparseCell).CompareTo
			Return Col.CompareTo(other.Col)
		End Function

		#End Region
	End Class
End Namespace
