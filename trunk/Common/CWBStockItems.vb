Imports Microsoft.Practices.EnterpriseLibrary.Common
Imports Microsoft.Practices.EnterpriseLibrary.Data
Imports System.Data
Imports System.Data.Common
Imports CWBCommon.CWBConstants
Imports DevExpress.XtraEditors
Public Class CWBStockItems
#Region "Variables"
    Private _StockID As Int64
    Private _Stock_Code As String
    Private _Description As String
    Private _Stock_Category As Int64
    Private _Purchasing_Price As Decimal
    Private _Sales_Price As Decimal
    Private _Supplier As Int64
    Private _Reorder_Level As Decimal
    Private _Stock_Balance As Decimal
    Private _Stock_Value As Decimal
    Private _StockCategoryName As String
    Private _StockCategoryID As Int64

#End Region

#Region "Properties"
    Public Property StockID() As Int64
        Get
            Return _StockID
        End Get
        Set(ByVal value As Int64)
            _StockID = value
        End Set
    End Property
    Public Property Stock_Code() As String
        Get
            Return _Stock_Code
        End Get
        Set(ByVal value As String)
            _Stock_Code = value
        End Set
    End Property
    Public Property Description() As String
        Get
            Return _Description
        End Get
        Set(ByVal value As String)
            _Description = value
        End Set
    End Property
    Public Property Stock_Category() As Int64
        Get
            Return _Stock_Category
        End Get
        Set(ByVal value As Int64)
            _Stock_Category = value
        End Set
    End Property
    Public Property Purchasing_Price() As Decimal
        Get
            Return _Purchasing_Price
        End Get
        Set(ByVal value As Decimal)
            _Purchasing_Price = value
        End Set
    End Property
    Public Property Sales_Price() As Decimal
        Get
            Return _Sales_Price
        End Get
        Set(ByVal value As Decimal)
            _Sales_Price = value
        End Set
    End Property
    Public Property Supplier() As Int64
        Get
            Return _Supplier
        End Get
        Set(ByVal value As Int64)
            _Supplier = value
        End Set
    End Property
    Public Property Reorder_Level() As Decimal
        Get
            Return _Reorder_Level
        End Get
        Set(ByVal value As Decimal)
            _Reorder_Level = value
        End Set
    End Property
    Public Property Stock_Balance() As Decimal
        Get
            Return _Stock_Balance
        End Get
        Set(ByVal value As Decimal)
            _Stock_Balance = value
        End Set
    End Property
    Public Property Stock_Value() As Decimal
        Get
            Return _Stock_Value
        End Get
        Set(ByVal value As Decimal)
            _Stock_Value = value
        End Set
    End Property
    Public Property StockCategoryName() As String
        Get
            Return _StockCategoryName
        End Get
        Set(ByVal value As String)
            _StockCategoryName = value
        End Set
    End Property
    Public Property StockCategoryID() As Int64
        Get
            Return _StockCategoryID
        End Get
        Set(ByVal value As Int64)
            _StockCategoryID = value
        End Set
    End Property
#End Region

#Region "Insert Stock"
    Public Sub InsertStock()

        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(STOCK_INSERT)
            DB.AddInParameter(DBC, "@StockID", DbType.Int64, Me.StockID)
            DB.AddInParameter(DBC, "@Stock_Code", DbType.String, Me.Stock_Code)
            DB.AddInParameter(DBC, "@Description", DbType.String, Me.Description)
            DB.AddInParameter(DBC, "@Stock_Category", DbType.Int64, Me.Stock_Category)
            DB.AddInParameter(DBC, "@Purchasing_Price", DbType.Decimal, Me.Purchasing_Price)
            DB.AddInParameter(DBC, "@Sales_Price", DbType.Decimal, Me.Sales_Price)
            DB.AddInParameter(DBC, "@Supplier", DbType.Int64, Me.Supplier)
            DB.AddInParameter(DBC, "@Reorder_Level", DbType.Decimal, Me.Reorder_Level)
            DB.AddInParameter(DBC, "@Stock_Balance", DbType.Decimal, Me.Stock_Balance)
            DB.AddInParameter(DBC, "@Stock_value", DbType.Decimal, Me.Stock_Value)
            DB.ExecuteNonQuery(DBC)
        Catch ex As Exception




        End Try

    End Sub

#End Region

#Region "Delete Stock"
    Public Sub DeleteStock()
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(STOCK_DELETE)
            DB.AddInParameter(DBC, "@StockID", DbType.Int64, Me.StockID)
            DB.ExecuteNonQuery(DBC)
        Catch ex As Exception

        End Try

    End Sub
#End Region

#Region "Get Stock By Stock ID"
    Public Sub GetByStockID()
        Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
        Dim DBC As DbCommand = DB.GetStoredProcCommand(STOCK_GETBYID)
        Try

            DB.AddInParameter(DBC, "@StockID", DbType.Int64, Me.StockID)

            Using DR As IDataReader = DB.ExecuteReader(DBC)


                With DR
                    Do While .Read
                        Me.Stock_Code = IIf(Not IsDBNull(.Item("StockCode")), Trim(.Item("StockCode").ToString), String.Empty)
                        Me.Description = IIf(Not IsDBNull(.Item("Description")), Trim(.Item("Description").ToString), String.Empty)
                        Me.Stock_Category = Convert.ToInt64(IIf(Not IsDBNull(.Item("StockCategory")), Trim(.Item("StockCategory").ToString), 0))
                        Me.Purchasing_Price = IIf(Not IsDBNull(.Item("PurchasingPrice")), Decimal.Parse(Trim(.Item("PurchasingPrice")).ToString), 0)
                        Me.Sales_Price = IIf(Not IsDBNull(.Item("SalesPrice")), Decimal.Parse(Trim(.Item("SalesPrice")).ToString), 0)
                        Me.Supplier = Convert.ToInt64(IIf(Not IsDBNull(.Item("Supplier")), Trim(.Item("Supplier").ToString), 0))
                        Me.Reorder_Level = IIf(Not IsDBNull(.Item("ReorderLevel")), Decimal.Parse(Trim(.Item("ReorderLevel")).ToString), 0)
                        Me.Stock_Balance = IIf(Not IsDBNull(.Item("StockBalance")), Decimal.Parse(Trim(.Item("StockBalance")).ToString), 0)
                        Me.Stock_Value = IIf(Not IsDBNull(.Item("StockValue")), Decimal.Parse(Trim(.Item("StockValue")).ToString), 0)

                    Loop
                End With

                If (Not DR Is Nothing) Then
                    DR.Close()
                End If


            End Using



        Catch ex As Exception

        Finally
            DBC.Dispose()



        End Try


    End Sub
#End Region

#Region "Get All StockItems"
    Public Function GetAllStockItems() As DataSet
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(STOCK_GETALL)
            Return DB.ExecuteDataSet(DBC)
            DBC.Dispose()
        Catch ex As Exception

            Return Nothing
        End Try




    End Function
#End Region

#Region "Insert Stock Category"
    Public Sub InsertStockCategory()

        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(STOCKCATEGORY_INSERT)
            DB.AddInParameter(DBC, "@Name", DbType.String, Me.StockCategoryName)
            DB.ExecuteNonQuery(DBC)
        Catch ex As Exception




        End Try

    End Sub

#End Region

#Region "Delete Stock Category"
    Public Sub DeleteStockCategories()
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(STOCKCATEGORY_DELETE)
            DB.AddInParameter(DBC, "@Name", DbType.String, Me.StockCategoryName)
            DB.ExecuteNonQuery(DBC)
        Catch ex As Exception

        End Try

    End Sub
#End Region

#Region "Get All Stock Categories"
    Public Function GetAllStockCategories() As DataSet
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(STOCKCATEGORY_GETALL)
            Return DB.ExecuteDataSet(DBC)
            DBC.Dispose()
        Catch ex As Exception

            Return Nothing
        End Try




    End Function
#End Region

End Class
