Imports Microsoft.Practices.EnterpriseLibrary.Common
Imports Microsoft.Practices.EnterpriseLibrary.Data
Imports System.Data
Imports System.Data.Common
Imports CWBCommon.CWBConstants
Imports DevExpress.XtraEditors
Public Class CWBStock

#Region "Variables"
    Private _StockID As Int64
    Private _StockCode As String
    Private _Description As String
    Private _StockCategoryID As Int64
    Private _StockType As String
    Private _PurchasingPrice As Decimal
    Private _SalesPrice As Decimal
    Private _SupplierID As Int64
    Private _ReorderLevel As Decimal
    Private _OpeningStockBalance As Decimal
    Private _OpeningStockValue As Decimal
    Private _CurrentStockBalance As Decimal
    Private _CurrentStockValue As Decimal
    Private _StockCategoryName As String
    Private _CreatedBy As Int64
    Private _CreatedDate As DateTime
    Private _UpdatedBy As Int64
    Private _UpdatedDate As DateTime


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
    Public Property StockCode() As String
        Get
            Return _StockCode
        End Get
        Set(ByVal value As String)
            _StockCode = value
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
    Public Property StockType() As String
        Get
            Return _StockType
        End Get
        Set(ByVal value As String)
            _StockType = value
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
    Public Property PurchasingPrice() As Decimal
        Get
            Return _PurchasingPrice
        End Get
        Set(ByVal value As Decimal)
            _PurchasingPrice = value
        End Set
    End Property
    Public Property SalesPrice() As Decimal
        Get
            Return _SalesPrice
        End Get
        Set(ByVal value As Decimal)
            _SalesPrice = value
        End Set
    End Property
    Public Property SupplierID() As Int64
        Get
            Return _SupplierID
        End Get
        Set(ByVal value As Int64)
            _SupplierID = value
        End Set
    End Property
    Public Property ReorderLevel() As Decimal
        Get
            Return _ReorderLevel
        End Get
        Set(ByVal value As Decimal)
            _ReorderLevel = value
        End Set
    End Property
    Public Property OpeningStockBalance() As Decimal
        Get
            Return _OpeningStockBalance
        End Get
        Set(ByVal value As Decimal)
            _OpeningStockBalance = value
        End Set
    End Property
    Public Property OpeningStockValue() As Decimal
        Get
            Return _OpeningStockValue
        End Get
        Set(ByVal value As Decimal)
            _OpeningStockValue = value
        End Set
    End Property
    Public Property CurrentStockBalance() As Decimal
        Get
            Return _CurrentStockBalance
        End Get
        Set(ByVal value As Decimal)
            _CurrentStockBalance = value
        End Set
    End Property
    Public Property CurrentStockValue() As Decimal
        Get
            Return _CurrentStockValue
        End Get
        Set(ByVal value As Decimal)
            _CurrentStockValue = value
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
    Public Property CreatedBy() As Int64
        Get
            Return _CreatedBy
        End Get
        Set(ByVal value As Int64)
            _CreatedBy = value
        End Set
    End Property
    Public Property CreatedDate() As DateTime
        Get
            Return _CreatedDate
        End Get
        Set(ByVal value As DateTime)
            _CreatedDate = value
        End Set
    End Property
    Public Property UpdatedBy() As Int64
        Get
            Return _UpdatedBy
        End Get
        Set(ByVal value As Int64)
            _UpdatedBy = value
        End Set
    End Property
    Public Property UpdatedDate() As DateTime
        Get
            Return _UpdatedDate
        End Get
        Set(ByVal value As DateTime)
            _UpdatedDate = value
        End Set
    End Property

#End Region

#Region "Insert Stock"
    Public Sub InsertStock()

        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(STOCK_INSERT)
            DB.AddInParameter(DBC, "@StockID", DbType.Int64, Me.StockID)
            DB.AddInParameter(DBC, "@StockCode", DbType.String, Me.StockCode)
            DB.AddInParameter(DBC, "@Description", DbType.String, Me.Description)
            DB.AddInParameter(DBC, "@StockType", DbType.String, Me.StockType)
            DB.AddInParameter(DBC, "@StockCategoryID", DbType.Int64, Me.StockCategoryID)
            DB.AddInParameter(DBC, "@PurchasingPrice", DbType.Decimal, Me.PurchasingPrice)
            DB.AddInParameter(DBC, "@SalesPrice", DbType.Decimal, Me.SalesPrice)
            DB.AddInParameter(DBC, "@SupplierID", DbType.Int64, Me.SupplierID)
            DB.AddInParameter(DBC, "@ReorderLevel", DbType.Decimal, Me.ReorderLevel)
            DB.AddInParameter(DBC, "@OpeningStockBalance", DbType.Decimal, Me.OpeningStockBalance)
            DB.AddInParameter(DBC, "@OpeningStockValue", DbType.Decimal, Me.OpeningStockValue)
            DB.AddInParameter(DBC, "@CurrentStockBalance", DbType.Decimal, Me.CurrentStockBalance)
            DB.AddInParameter(DBC, "@CurrentStockValue", DbType.Decimal, Me.CurrentStockValue)
            DB.AddInParameter(DBC, "@CreatedBy", DbType.Int64, Me.CreatedBy)
            DB.AddInParameter(DBC, "@UpdatedBy", DbType.Int64, Me.UpdatedBy)
            DB.ExecuteNonQuery(DBC)
        Catch ex As Exception

            Throw



        End Try

    End Sub

#End Region

#Region "Is Stock Exits"
    Public Function IsStockExists() As Boolean
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(STOCK_ISSTOCKEXISTS)

            DB.AddInParameter(DBC, "@StockCode", DbType.String, StockCode)
            DB.AddOutParameter(DBC, "@IsExists", DbType.Int64, 0)
            DB.ExecuteNonQuery(DBC)

            If DB.GetParameterValue(DBC, "@IsExists") = 1 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
            Throw

        End Try

    End Function
#End Region

#Region "Delete Stock"
    Public Sub DeleteStock()
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(STOCK_DELETE)
            DB.AddInParameter(DBC, "@StockID", DbType.Int64, Me.StockID)
            DB.ExecuteNonQuery(DBC)
        Catch ex As Exception
            Throw

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
                        Me.StockID = Convert.ToInt64(IIf(Not IsDBNull(.Item("StockID")), Trim(.Item("StockID").ToString), 0))
                        Me.StockCode = IIf(Not IsDBNull(.Item("StockCode")), Trim(.Item("StockCode").ToString), String.Empty)
                        Me.Description = IIf(Not IsDBNull(.Item("Description")), Trim(.Item("Description").ToString), String.Empty)
                        Me.StockType = IIf(Not IsDBNull(.Item("StockType")), Trim(.Item("StockType").ToString), String.Empty)

                        Me.StockCategoryID = Convert.ToInt64(IIf(Not IsDBNull(.Item("StockCategoryID")), Trim(.Item("StockCategoryID").ToString), 0))
                        Me.PurchasingPrice = IIf(Not IsDBNull(.Item("PurchasingPrice")), Decimal.Parse(Trim(.Item("PurchasingPrice")).ToString), 0)
                        Me.SalesPrice = IIf(Not IsDBNull(.Item("SalesPrice")), Decimal.Parse(Trim(.Item("SalesPrice")).ToString), 0)
                        Me.SupplierID = Convert.ToInt64(IIf(Not IsDBNull(.Item("SupplierID")), Trim(.Item("SupplierID").ToString), 0))
                        Me.ReorderLevel = IIf(Not IsDBNull(.Item("ReorderLevel")), Decimal.Parse(Trim(.Item("ReorderLevel")).ToString), 0)
                        Me.OpeningStockBalance = Convert.ToDecimal(IIf(Not IsDBNull(.Item("OpeningStockBalance")), Decimal.Parse(Trim(.Item("OpeningStockBalance")).ToString), 0))
                        Me.OpeningStockValue = Convert.ToDecimal(IIf(Not IsDBNull(.Item("OpeningStockValue")), Decimal.Parse(Trim(.Item("OpeningStockValue")).ToString), 0))
                        Me.CurrentStockBalance = Convert.ToDecimal(IIf(Not IsDBNull(.Item("CurrentStockBalance")), Decimal.Parse(Trim(.Item("CurrentStockBalance")).ToString), 0))
                        Me.CurrentStockValue = Convert.ToDecimal(IIf(Not IsDBNull(.Item("CurrentStockValue")), Decimal.Parse(Trim(.Item("CurrentStockValue")).ToString), 0))
                        Me.CreatedBy = Convert.ToInt64(IIf(Not IsDBNull(.Item("CreatedBy")), Trim(.Item("CreatedBy").ToString), 0))
                        Me.UpdatedBy = Convert.ToInt64(IIf(Not IsDBNull(.Item("UpdatedBy")), Trim(.Item("UpdatedBy").ToString), 0))

                    Loop
                End With

                If (Not DR Is Nothing) Then
                    DR.Close()
                End If


            End Using



        Catch ex As Exception
            Throw

        Finally
            DBC.Dispose()



        End Try


    End Sub
#End Region

#Region "Get Stock By StockCode"
    Public Sub GetByStockCode()
        Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
        Dim DBC As DbCommand = DB.GetStoredProcCommand(STOCK_GETBYSTOCKCODE)
        Try

            DB.AddInParameter(DBC, "@StockCode", DbType.String, Me.StockCode)

            Using DR As IDataReader = DB.ExecuteReader(DBC)


                With DR
                    Do While .Read
                        Me.StockID = Convert.ToInt64(IIf(Not IsDBNull(.Item("StockID")), Trim(.Item("StockID").ToString), 0))
                        Me.StockCode = IIf(Not IsDBNull(.Item("StockCode")), Trim(.Item("StockCode").ToString), String.Empty)
                        Me.Description = IIf(Not IsDBNull(.Item("Description")), Trim(.Item("Description").ToString), String.Empty)
                        Me.StockType = IIf(Not IsDBNull(.Item("StockType")), Trim(.Item("StockType").ToString), String.Empty)

                        Me.StockCategoryID = Convert.ToInt64(IIf(Not IsDBNull(.Item("StockCategoryID")), Trim(.Item("StockCategoryID").ToString), 0))
                        Me.PurchasingPrice = IIf(Not IsDBNull(.Item("PurchasingPrice")), Decimal.Parse(Trim(.Item("PurchasingPrice")).ToString), 0)
                        Me.SalesPrice = IIf(Not IsDBNull(.Item("SalesPrice")), Decimal.Parse(Trim(.Item("SalesPrice")).ToString), 0)
                        Me.SupplierID = Convert.ToInt64(IIf(Not IsDBNull(.Item("SupplierID")), Trim(.Item("SupplierID").ToString), 0))
                        Me.ReorderLevel = IIf(Not IsDBNull(.Item("ReorderLevel")), Decimal.Parse(Trim(.Item("ReorderLevel")).ToString), 0)
                        Me.OpeningStockBalance = Convert.ToDecimal(IIf(Not IsDBNull(.Item("OpeningStockBalance")), Decimal.Parse(Trim(.Item("OpeningStockBalance")).ToString), 0))
                        Me.OpeningStockValue = Convert.ToDecimal(IIf(Not IsDBNull(.Item("OpeningStockValue")), Decimal.Parse(Trim(.Item("OpeningStockValue")).ToString), 0))
                        Me.CurrentStockBalance = Convert.ToDecimal(IIf(Not IsDBNull(.Item("CurrentStockBalance")), Decimal.Parse(Trim(.Item("CurrentStockBalance")).ToString), 0))
                        Me.CurrentStockValue = Convert.ToDecimal(IIf(Not IsDBNull(.Item("CurrentStockValue")), Decimal.Parse(Trim(.Item("CurrentStockValue")).ToString), 0))
                        Me.CreatedBy = Convert.ToInt64(IIf(Not IsDBNull(.Item("CreatedBy")), Trim(.Item("CreatedBy").ToString), 0))
                        Me.UpdatedBy = Convert.ToInt64(IIf(Not IsDBNull(.Item("UpdatedBy")), Trim(.Item("UpdatedBy").ToString), 0))

                    Loop
                End With

                If (Not DR Is Nothing) Then
                    DR.Close()
                End If


            End Using



        Catch ex As Exception
            Throw

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
            Throw

        End Try




    End Function
#End Region

#Region "Get All StockItems to Order"
    Public Function GetAllStockItemsToOrder() As DataSet
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(STOCK_ITEMSTOORDER)
            Return DB.ExecuteDataSet(DBC)
            DBC.Dispose()
        Catch ex As Exception

            Return Nothing
            Throw

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

            Throw



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
            Throw

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
            Throw

        End Try




    End Function
#End Region

End Class
