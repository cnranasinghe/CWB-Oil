Imports Microsoft.Practices.EnterpriseLibrary.Common
Imports Microsoft.Practices.EnterpriseLibrary.Data
Imports System.Data
Imports System.Data.Common
Imports CWBCommon.CWBConstants
Imports DevExpress.XtraEditors
Imports DevExpress.XtraGrid
Imports System
Public Class CWBSales

#Region "Variables"
    Private _SalesID As Int64
    Private _CurrentSalesID As Int64
    Private _StockID As Int64
    Private _CustomerID As Int64
    Private _ReferenceNo As String
    Private _SalesDate As DateTime
    Private _Total As Decimal
    Private _TaxPercent As Decimal
    Private _TaxAmount As Decimal
    Private _Discount As Decimal
    Private _Note As String
    Private _Paid As Boolean

    Private _UnitPrice As Decimal
    Private _PurchasingPrice As Decimal
    Private _Quantity As Decimal
    Private _PDiscount As Decimal
    Private _Value As Decimal
    Private _StockBalance As Decimal

#End Region

#Region "Properties"
    Public Property SalesID() As Int64
        Get
            Return _SalesID
        End Get
        Set(ByVal value As Int64)
            _SalesID = value
        End Set
    End Property
    Public Property CurrentSalesID() As Int64
        Get
            Return _CurrentSalesID
        End Get
        Set(ByVal value As Int64)
            _CurrentSalesID = value
        End Set
    End Property
    Public Property StockID() As Int64
        Get
            Return _StockID
        End Get
        Set(ByVal value As Int64)
            _StockID = value
        End Set
    End Property
    Public Property CustomerID() As Int64
        Get
            Return _CustomerID
        End Get
        Set(ByVal value As Int64)
            _CustomerID = value
        End Set
    End Property
    Public Property ReferenceNo() As String
        Get
            Return _ReferenceNo
        End Get
        Set(ByVal value As String)
            _ReferenceNo = value
        End Set
    End Property
    Public Property SalesDate() As DateTime
        Get
            Return _SalesDate
        End Get
        Set(ByVal value As DateTime)
            _SalesDate = value
        End Set
    End Property
    Public Property Total() As Decimal
        Get
            Return _Total
        End Get
        Set(ByVal value As Decimal)
            _Total = value
        End Set
    End Property
    Public Property TaxPercent() As Decimal
        Get
            Return _TaxPercent
        End Get
        Set(ByVal value As Decimal)
            _TaxPercent = value
        End Set
    End Property
    Public Property TaxAmount() As Decimal
        Get
            Return _TaxAmount
        End Get
        Set(ByVal value As Decimal)
            _TaxAmount = value
        End Set
    End Property
    Public Property Discount() As Decimal
        Get
            Return _Discount
        End Get
        Set(ByVal value As Decimal)
            _Discount = value
        End Set
    End Property
    Public Property Note() As String
        Get
            Return _Note
        End Get
        Set(ByVal value As String)
            _Note = value
        End Set
    End Property
    Public Property Paid() As Boolean
        Get
            Return _Paid
        End Get
        Set(ByVal value As Boolean)
            _Paid = value
        End Set
    End Property
    Public Property UnitPrice() As Decimal
        Get
            Return _UnitPrice
        End Get
        Set(ByVal value As Decimal)
            _UnitPrice = value
        End Set
    End Property
    Public Property Quantity() As Decimal
        Get
            Return _Quantity
        End Get
        Set(ByVal value As Decimal)
            _Quantity = value
        End Set
    End Property
    Public Property PurchasingPrice() As Decimal
        Get
            Return _PurchasingPrice
        End Get
        Set(ByVal value As Decimal)
            _PurchasingPrice = Value
        End Set
    End Property
    Public Property PDiscount() As Decimal
        Get
            Return _PDiscount
        End Get
        Set(ByVal value As Decimal)
            _PDiscount = value
        End Set
    End Property
    Public Property Value() As Decimal
        Get
            Return _Value
        End Get
        Set(ByVal value As Decimal)
            _Value = Value
        End Set
    End Property
    Public Property StockBalance() As Decimal
        Get
            Return _StockBalance
        End Get
        Set(ByVal value As Decimal)
            _StockBalance = Value
        End Set
    End Property
#End Region

#Region "Insert Sales"
    Public Sub InsertSales(ByVal db As Database, ByVal transaction As DbTransaction)

        Try

            Dim DBC As DbCommand = db.GetStoredProcCommand(SALES_INSERT)
            db.AddInParameter(DBC, "@SalesID", DbType.Int64, Me.SalesID)
            db.AddInParameter(DBC, "@CustomerID", DbType.Int64, Me.CustomerID)
            db.AddInParameter(DBC, "@ReferenceNo", DbType.String, Me.ReferenceNo)
            db.AddInParameter(DBC, "@SalesDate", DbType.DateTime, Me.SalesDate)
            db.AddInParameter(DBC, "@Total", DbType.Decimal, Me.Total)
            db.AddInParameter(DBC, "@Percentage", DbType.Decimal, Me.TaxPercent)
            db.AddInParameter(DBC, "@Tax", DbType.Decimal, Me.TaxAmount)
            db.AddInParameter(DBC, "@Discount", DbType.Decimal, Me.Discount)
            db.AddInParameter(DBC, "@Note", DbType.String, Me.Note)
            db.AddInParameter(DBC, "@Paid", DbType.Boolean, Me.Paid)
            db.AddOutParameter(DBC, "@CurrentSalesID", DbType.Int64, 0)

            db.ExecuteNonQuery(DBC, transaction)

            Me.CurrentSalesID = Convert.ToInt64(db.GetParameterValue(DBC, "@CurrentSalesID").ToString())
        Catch ex As Exception


        End Try

    End Sub
#End Region

#Region "Insert Sales Description"
    Public Sub InsertSalesDescription(ByVal db As Database, ByVal transaction As DbTransaction)

        Try
            'Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = db.GetStoredProcCommand(SALESDESCRIPTION_INSERT)
            DB.AddInParameter(DBC, "@SalesID", DbType.Int64, Me.SalesID)
            DB.AddInParameter(DBC, "@StockID", DbType.Int64, Me.StockID)
            DB.AddInParameter(DBC, "@UnitPrice", DbType.Decimal, Me.UnitPrice)
            DB.AddInParameter(DBC, "@PurchasingPrice", DbType.Decimal, Me.PurchasingPrice)
            DB.AddInParameter(DBC, "@Quantity", DbType.Decimal, Me.Quantity)
            DB.AddInParameter(DBC, "@Discount", DbType.Decimal, Me.PDiscount)
            DB.AddInParameter(DBC, "@Value", DbType.Decimal, Me.Value)
            db.ExecuteNonQuery(DBC, transaction)

        Catch ex As Exception


        End Try

    End Sub
#End Region

#Region "Get Purchas By SalesID"
    Public Sub GetSalesByID()
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(SALES_GETBYID)
            DB.AddInParameter(DBC, "@SalesID", DbType.Int64, Me.SalesID)

            Using DR As IDataReader = DB.ExecuteReader(DBC)


                With DR
                    Do While .Read

                        Me.CustomerID = Convert.ToInt64(IIf(Not IsDBNull(.Item("CustomerID")), Trim(.Item("CustomerID").ToString), 0))
                        Me.ReferenceNo = IIf(Not IsDBNull(.Item("ReferenceNo")), Trim(.Item("ReferenceNo").ToString), String.Empty)
                        Me.SalesDate = IIf(Not IsDBNull(.Item("SalesDate")), DateTime.Parse(Trim(.Item("SalesDate").ToString)), String.Empty)
                        Me.Total = Convert.ToDecimal(IIf(Not IsDBNull(.Item("Total")), .Item("Total").ToString, 0))
                        Me.TaxPercent = Convert.ToDecimal(IIf(Not IsDBNull(.Item("Percentage")), .Item("Percentage").ToString, 0))
                        Me.TaxAmount = Convert.ToDecimal(IIf(Not IsDBNull(.Item("Tax")), .Item("Tax").ToString, String.Empty))
                        Me.Discount = Convert.ToDecimal(IIf(Not IsDBNull(.Item("Discount")), .Item("Discount").ToString, 0))
                        Me.Note = IIf(Not IsDBNull(.Item("Note")), Trim(.Item("Note").ToString), String.Empty)
                        Me.Paid = IIf(Not IsDBNull(.Item("Paid")), Boolean.Parse(Trim(.Item("Paid").ToString)), String.Empty)


                    Loop
                End With

                If (Not DR Is Nothing) Then
                    DR.Close()
                End If


            End Using

        Catch ex As Exception

        End Try
    End Sub
#End Region

#Region "Get SalesDescription By SalesID"
    Public Function GetSalesDescriptionByID() As DataSet
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(SALESDESCRIPTION_GETBYID)
            DB.AddInParameter(DBC, "@SalesID", DbType.Int64, Me.SalesID)
            Return DB.ExecuteDataSet(DBC)

        Catch ex As Exception
            Return Nothing
        End Try
    End Function
#End Region

#Region "GetAll Sales"
    Public Function GetAllSales() As DataSet
        Try

            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(SALES_GETALL)
            Return DB.ExecuteDataSet(DBC)

        Catch ex As Exception
            Return Nothing
        End Try

    End Function
#End Region

#Region "Sales Description Delete"
    Public Sub SalesDescriptionDelete(ByVal db As Database, ByVal transaction As DbTransaction)

        Try
            'Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = db.GetStoredProcCommand(SALESDESCRIPTION_DELETE)

            DB.AddInParameter(DBC, "@SalesID", DbType.Int64, Me.SalesID)
            db.ExecuteNonQuery(DBC, transaction)

        Catch ex As Exception


        End Try

    End Sub
#End Region

#Region "Add to Stock from Sales"
    Public Sub AddToStock(ByVal db As Database, ByVal transaction As DbTransaction)

        Try

            Dim DBC As DbCommand = db.GetStoredProcCommand(STOCK_ADDROMSALES)
            DB.AddInParameter(DBC, "@SalesID", DbType.Int64, Me.SalesID)
            db.ExecuteNonQuery(DBC, transaction)

        Catch ex As Exception


        End Try

    End Sub
#End Region

#Region "Update Stock from Sales"
    Public Sub UpdateStock(ByVal db As Database, ByVal transaction As DbTransaction)

        Try
            'Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = db.GetStoredProcCommand(STOCK_UPDATEBYSALES)

            DB.AddInParameter(DBC, "@StockID", DbType.Int64, Me.StockID)
            DB.AddInParameter(DBC, "@StockBalance", DbType.Decimal, Me.StockBalance)
            db.ExecuteNonQuery(DBC, transaction)

        Catch ex As Exception


        End Try

    End Sub
#End Region

#Region "Get All Unsettled Sales Bills"
    Public Function GetAllSalesUnsettledBills() As DataSet
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(SALES_UNSETTLED_BILLS)
            Return DB.ExecuteDataSet(DBC)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
#End Region

#Region "Get All Settled Sales Bills"
    Public Function GetAllSalesSettledBills() As DataSet
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(SALES_SETTLED_BILLS)
            Return DB.ExecuteDataSet(DBC)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
#End Region
End Class
