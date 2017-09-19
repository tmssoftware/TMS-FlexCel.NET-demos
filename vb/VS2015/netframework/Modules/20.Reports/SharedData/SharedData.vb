Imports System
Imports System.Data
Imports System.Reflection
Imports FlexCel.Report
Imports System.IO

Namespace FlexCel.Demo.SharedData
	''' <summary>
	''' A common interface for all demos using Northwind database.
	''' Being centralized here, you can replace the data access just here and it will change for all the demos.
	'''<br></br>
	''' This class uses a lot of awful hacks to retrieve a dataset, and it also loads all tables even if you need a couple.
	''' We had to do it this way to keep examples simple, and to avoid having references to SQLite, which you might not have installed.
	''' In a real application, you should use a different approach, only load the tables you need, etc. In the demos here, that performance doesn't really matter,
	''' and the demos are to show you how to create reports, not how to fill datasets. We assume you already know how to fill a dataset.
	''' </summary>
	Public Class SharedData
		Private Shared Function DataPath() As String
			Return Path.Combine(Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), ".."), "..") & Path.DirectorySeparatorChar
		End Function

		Private Shared Function GetSQLAssembly() As System.Reflection.Assembly
			Dim SQLitePath As String = Path.Combine(DataPath(), "..\SharedData\SQLite")

			If IntPtr.Size = 8 Then ' 64 bit.
				Return System.Reflection.Assembly.LoadFile(Path.Combine(SQLitePath, "x64\System.Data.SQLite.dll"))
			Else '32 bit.
				Return System.Reflection.Assembly.LoadFile(Path.Combine(SQLitePath, "x86\System.Data.SQLite.dll"))
			End If
		End Function

		Private Shared Function GetConnection() As IDbConnection
			'In this demo we are going to load the assembly dynamically, so you can run other demos even if there are any issues with SQLite.
			'Normally you would just add a reference to System.Data.SQLite or use the SQLite nuget package.

			Dim asm As System.Reflection.Assembly = GetSQLAssembly()
			Dim ConnectionType As Type = asm.GetType("System.Data.SQLite.SQLiteConnection")
			Dim Connection As IDbConnection = CType(Activator.CreateInstance(ConnectionType), IDbConnection)

			Dim DbPath As String = Path.Combine(DataPath(), "..\SharedData\Northwind.sqlite")
			Connection.ConnectionString = String.Format("Data Source={0}", DbPath)
			Return Connection
		End Function

		Private Shared Function GetDataAdapter(ByVal Connection As IDbConnection, ByVal Sql As String) As IDbDataAdapter
			Dim asm As System.Reflection.Assembly = GetSQLAssembly()
			Dim DataAdapterType As Type = asm.GetType("System.Data.SQLite.SQLiteDataAdapter")
			Dim DataAdapter As IDbDataAdapter = CType(Activator.CreateInstance(DataAdapterType, Sql, Connection), IDbDataAdapter)
			Return DataAdapter
		End Function

		Private Shared Function GetDataTable() As DataSet
			Using Connection As IDbConnection = GetConnection()
				Dim ds As New DataSet()
				ds.EnforceConstraints = False

				AddTable(Connection, ds, "SELECT * FROM Suppliers", "Suppliers")
				AddTable(Connection, ds, "SELECT * FROM Categories", "Categories")

				AddTable(Connection, ds, "SELECT * FROM Products", "Products")

				AddTable(Connection, ds, "SELECT * FROM Customers", "Customers")
				AddTable(Connection, ds, "SELECT * FROM Shippers", "Shippers")
				AddTable(Connection, ds, "SELECT * FROM Employees", "Employees")

				AddTable(Connection, ds, "SELECT * FROM Orders", "Orders")
				AddTable(Connection, ds, "SELECT * FROM [Order Details]", "Order Details")

				AddRelationships(ds)
				ds.EnforceConstraints = True

				Return ds

			End Using
		End Function

		Private Shared Sub AddRelationships(ByVal ds As DataSet)
			Dim relationOrder_Details_FK00 As New DataRelation("Order Details_FK00", New DataColumn() { ds.Tables("Products").Columns("ProductID")}, New DataColumn() { ds.Tables("Order Details").Columns("ProductID")}, False)
			ds.Relations.Add(relationOrder_Details_FK00)

			Dim relationOrder_Details_FK01 As New DataRelation("Order Details_FK01", New DataColumn() { ds.Tables("Orders").Columns("OrderID")}, New DataColumn() { ds.Tables("Order Details").Columns("OrderID")}, False)
			ds.Relations.Add(relationOrder_Details_FK01)

			Dim relationOrders_FK00 As New DataRelation("Orders_FK00", New DataColumn() { ds.Tables("Customers").Columns("CustomerID")}, New DataColumn() { ds.Tables("Orders").Columns("CustomerID")}, False)
			ds.Relations.Add(relationOrders_FK00)

			Dim relationOrders_FK01 As New DataRelation("Orders_FK01", New DataColumn() { ds.Tables("Shippers").Columns("ShipperID")}, New DataColumn() { ds.Tables("Orders").Columns("ShipVia")}, False)
			ds.Relations.Add(relationOrders_FK01)

			Dim relationOrders_FK02 As New DataRelation("Orders_FK02", New DataColumn() { ds.Tables("Employees").Columns("EmployeeID")}, New DataColumn() { ds.Tables("Orders").Columns("EmployeeID")}, False)
			ds.Relations.Add(relationOrders_FK02)

			Dim relationProducts_FK00 As New DataRelation("Products_FK00", New DataColumn() { ds.Tables("Suppliers").Columns("SupplierID")}, New DataColumn() { ds.Tables("Products").Columns("SupplierID")}, False)
			ds.Relations.Add(relationProducts_FK00)

			Dim relationProducts_FK01 As New DataRelation("Products_FK01", New DataColumn() { ds.Tables("Categories").Columns("CategoryID")}, New DataColumn() { ds.Tables("Products").Columns("CategoryID")}, False)
			ds.Relations.Add(relationProducts_FK01)
		End Sub

		Private Shared Sub AddTable(ByVal Connection As IDbConnection, ByVal ds As DataSet, ByVal Sql As String, ByVal TableName As String)
			Dim DataAdapter As IDbDataAdapter = GetDataAdapter(Connection, Sql)
			Try
				DataAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey
				DataAdapter.Fill(ds)
				ds.Tables("Table").TableName = TableName
			Finally
				CType(DataAdapter, IDisposable).Dispose()
			End Try
		End Sub

		Friend Shared Function CreateReport() As FlexCel.Report.FlexCelReport
			Dim Result As New FlexCelReport(True)
			Result.AddTable(GetDataTable())
			Return Result
		End Function

		Friend Shared Sub Fill(ByVal ds As DataSet, ByVal sql As String, ByVal TableName As String)
			Using Connection As IDbConnection = GetConnection()
				AddTable(Connection, ds, sql, TableName)
			End Using
		End Sub

		Friend Shared Function GetDataAdapter() As IDbDataAdapter
			Dim Connection As IDbConnection = GetConnection()
			Return GetDataAdapter(Connection, "")
		End Function

		Friend Shared Function GetOrders() As DataTable
			Using Connection As IDbConnection = GetConnection()
				Dim ds As New DataSet()
				AddTable(Connection, ds, "select * from orders", "Orders")
				Return ds.Tables("Orders")
			End Using
		End Function

		Friend Shared Function CreateParameter(ByVal ParamName As String, ByVal ParamValue As Object) As IDbDataParameter
			Dim asm As System.Reflection.Assembly = GetSQLAssembly()


			Dim ParameterType As Type = asm.GetType("System.Data.SQLite.SQLiteParameter")
			Return CType(Activator.CreateInstance(ParameterType, ParamName, ParamValue), IDbDataParameter)
		End Function
	End Class
End Namespace
