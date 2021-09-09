using System;
using System.Data;
using System.Reflection;
using FlexCel.Report;
using System.IO;

namespace FlexCel.Demo.SharedData
{
    /// <summary>
    /// A common interface for all demos using Northwind database.
    /// Being centralized here, you can replace the data access just here and it will change for all the demos.
    ///<br></br>
    /// This class uses a lot of awful hacks to retrieve a dataset, and it also loads all tables even if you need a couple.
    /// We had to do it this way to keep examples simple, and to avoid having references to SQLite, which you might not have installed.
    /// In a real application, you should use a different approach, only load the tables you need, etc. In the demos here, that performance doesn't really matter,
    /// and the demos are to show you how to create reports, not how to fill datasets. We assume you already know how to fill a dataset.
    /// </summary>
    public class SharedData
    {
        private static string DataPath()
        {
            return Path.Combine(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), ".."), "..") + Path.DirectorySeparatorChar;
        }

        private static Assembly GetSQLAssembly()
        {
            string SQLitePath = Path.Combine(DataPath(), @"..\SharedData\SQLite");

            if (IntPtr.Size == 8) // 64 bit.
            {
                return Assembly.LoadFile(Path.Combine(SQLitePath, @"x64\System.Data.SQLite.dll"));
            }
            else //32 bit.
            {
                return Assembly.LoadFile(Path.Combine(SQLitePath, @"x86\System.Data.SQLite.dll"));
            }
        }

        private static IDbConnection GetConnection()
        {
            //In this demo we are going to load the assembly dynamically, so you can run other demos even if there are any issues with SQLite.
            //Normally you would just add a reference to System.Data.SQLite or use the SQLite nuget package.

            Assembly asm = GetSQLAssembly();
            Type ConnectionType = asm.GetType("System.Data.SQLite.SQLiteConnection");
            IDbConnection Connection = (IDbConnection)Activator.CreateInstance(ConnectionType);

            string DbPath = Path.Combine(DataPath(), @"..\SharedData\Northwind.sqlite");
            Connection.ConnectionString = String.Format(@"Data Source={0}", DbPath);
            return Connection;
        }

        private static IDbDataAdapter GetDataAdapter(IDbConnection Connection, string Sql)
        {
            Assembly asm = GetSQLAssembly();
            Type DataAdapterType = asm.GetType("System.Data.SQLite.SQLiteDataAdapter");
            IDbDataAdapter DataAdapter = (IDbDataAdapter)Activator.CreateInstance(DataAdapterType, Sql, Connection);
            return DataAdapter;
        }

        private static DataSet GetDataTable()
        {
            using (IDbConnection Connection = GetConnection())
            {
                DataSet ds = new DataSet();
                ds.EnforceConstraints = false;

                AddTable(Connection, ds, "SELECT * FROM Suppliers", "Suppliers");
                AddTable(Connection, ds, "SELECT * FROM Categories", "Categories");

                AddTable(Connection, ds, "SELECT * FROM Products", "Products");

                AddTable(Connection, ds, "SELECT * FROM Customers", "Customers");
                AddTable(Connection, ds, "SELECT * FROM Shippers", "Shippers");
                AddTable(Connection, ds, "SELECT * FROM Employees", "Employees");

                AddTable(Connection, ds, "SELECT * FROM Orders", "Orders");
                AddTable(Connection, ds, "SELECT * FROM [Order Details]", "Order Details");

                AddRelationships(ds);
                ds.EnforceConstraints = true;

                return ds;

            }
        }

        private static void AddRelationships(DataSet ds)
        {
            DataRelation relationOrder_Details_FK00 = new DataRelation("Order Details_FK00", new DataColumn[] {
                        ds.Tables["Products"].Columns["ProductID"]}, new DataColumn[] {
                        ds.Tables["Order Details"].Columns["ProductID"]}, false);
            ds.Relations.Add(relationOrder_Details_FK00);

            DataRelation relationOrder_Details_FK01 = new DataRelation("Order Details_FK01", new DataColumn[] {
                        ds.Tables["Orders"].Columns["OrderID"]}, new DataColumn[] {
                        ds.Tables["Order Details"].Columns["OrderID"]}, false);
            ds.Relations.Add(relationOrder_Details_FK01);

            DataRelation relationOrders_FK00 = new DataRelation("Orders_FK00", new DataColumn[] {
                        ds.Tables["Customers"].Columns["CustomerID"]}, new DataColumn[] {
                        ds.Tables["Orders"].Columns["CustomerID"]}, false);
            ds.Relations.Add(relationOrders_FK00);

            DataRelation relationOrders_FK01 = new DataRelation("Orders_FK01", new DataColumn[] {
                        ds.Tables["Shippers"].Columns["ShipperID"]}, new DataColumn[] {
                        ds.Tables["Orders"].Columns["ShipVia"]}, false);
            ds.Relations.Add(relationOrders_FK01);

            DataRelation relationOrders_FK02 = new DataRelation("Orders_FK02", new DataColumn[] {
                        ds.Tables["Employees"].Columns["EmployeeID"]}, new DataColumn[] {
                        ds.Tables["Orders"].Columns["EmployeeID"]}, false);
            ds.Relations.Add(relationOrders_FK02);

            DataRelation relationProducts_FK00 = new DataRelation("Products_FK00", new DataColumn[] {
                        ds.Tables["Suppliers"].Columns["SupplierID"]}, new DataColumn[] {
                        ds.Tables["Products"].Columns["SupplierID"]}, false);
            ds.Relations.Add(relationProducts_FK00);

            DataRelation relationProducts_FK01 = new DataRelation("Products_FK01", new DataColumn[] {
                        ds.Tables["Categories"].Columns["CategoryID"]}, new DataColumn[] {
                        ds.Tables["Products"].Columns["CategoryID"]}, false);
            ds.Relations.Add(relationProducts_FK01);
        }

        private static void AddTable(IDbConnection Connection, DataSet ds, string Sql, string TableName)
        {
            IDbDataAdapter DataAdapter = GetDataAdapter(Connection, Sql);
            try
            {
                DataAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey;
                DataAdapter.Fill(ds);
                ds.Tables["Table"].TableName = TableName;
            }
            finally
            {
                ((IDisposable)DataAdapter).Dispose();
            }
        }

        internal static FlexCel.Report.FlexCelReport CreateReport()
        {
            FlexCelReport Result = new FlexCelReport(true);
            Result.AddTable(GetDataTable());
            return Result;
        }

        internal static void Fill(DataSet ds, string sql, string TableName)
        {
            using (IDbConnection Connection = GetConnection())
            {
                AddTable(Connection, ds, sql, TableName);
            }
        }

        internal static IDbDataAdapter GetDataAdapter()
        {
            IDbConnection Connection = GetConnection();
            return GetDataAdapter(Connection, "");
        }

        internal static DataTable GetOrders()
        {
            using (IDbConnection Connection = GetConnection())
            {
                DataSet ds = new DataSet();
                AddTable(Connection, ds, "select * from orders", "Orders");
                return ds.Tables["Orders"];
            }
        }

        internal static IDbDataParameter CreateParameter(string ParamName, object ParamValue)
        {
            Assembly asm = GetSQLAssembly();


            Type ParameterType = asm.GetType("System.Data.SQLite.SQLiteParameter");
            return (IDbDataParameter)Activator.CreateInstance(ParameterType, ParamName, ParamValue);
        }
    }
}
