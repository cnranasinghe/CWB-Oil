Imports Microsoft.Practices.EnterpriseLibrary.Common
Imports Microsoft.Practices.EnterpriseLibrary.Data
Imports System.Data
Imports System.Data.Common
Imports CWBCommon.CWBConstants
Imports DevExpress.XtraEditors
Imports DevExpress.XtraGrid
Imports System

Public Class CWBPurchases

#Region "Variables"

    Private _PurchaseID As Int64
    Private _CurrentPurchaseID As Int64

    Private _StockID As Int64

    Private _SupplierID As Int64
    Private _PurchaseNo As Int64
    Private _RefBillNo As String
    Private _PurchaseDate As DateTime
    Private _Total As Decimal
    Private _TaxPercent As Decimal
    Private _TaxAmount As Decimal
    Private _Discount As Decimal
    Private _Note As String
    Private _Paid As Boolean
    Private _StockCode As String
    Private _UnitPrice As Decimal
    Private _Quantity As Decimal
    Private _GrandTotal As Decimal
    Private _PDiscount As Decimal
    Private _Value As Decimal
    Private _PaymentType As String
    Private _CollectionDate As DateTime
    Private _Reference As String
    Private _CollectionAmount As Decimal
    Private _Type As String
    Private _StockBalance As Decimal
    Private _CreatedBy As Int64
    Private _CreatedDate As DateTime
    Private _UpdatedBy As Int64
    Private _UpdatedDate As DateTime
    Private _FromDate As DateTime
    Private _ToDate As DateTime

#End Region

#Region "Properties"

    Public Property PurchaseID() As Int64
        Get
            Return _PurchaseID
        End Get
        Set(ByVal value As Int64)
            _PurchaseID = value
        End Set
    End Property
    Public Property PurchaseNo() As Int64
        Get
            Return _PurchaseNo
        End Get
        Set(ByVal value As Int64)
            _PurchaseNo = value
        End Set
    End Property
    Public Property CurrentPurchaseID() As Int64
        Get
            Return _CurrentPurchaseID
        End Get
        Set(ByVal value As Int64)
            _CurrentPurchaseID = value
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
    Public Property SupplierID() As Int64
        Get
            Return _SupplierID
        End Get
        Set(ByVal value As Int64)
            _SupplierID = value
        End Set
    End Property
    Public Property RefBillNo() As String
        Get
            Return _RefBillNo
        End Get
        Set(ByVal value As String)
            _RefBillNo = value
        End Set
    End Property
    Public Property PurchaseDate() As DateTime
        Get
            Return _PurchaseDate
        End Get
        Set(ByVal value As DateTime)
            If _PurchaseDate = value Then
                Return
            End If
            _PurchaseDate = value
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
            If _TaxAmount = value Then
                Return
            End If
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
    Public Property GrandTotal() As Decimal
        Get
            Return _GrandTotal
        End Get
        Set(ByVal value As Decimal)
            _GrandTotal = value
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
    Public Property Stock_Code() As String
        Get
            Return _StockCode
        End Get
        Set(ByVal value As String)
            _StockCode = value
        End Set
    End Property
    Public Property Unit_Price() As Decimal
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
            _Value = value
        End Set
    End Property
    Public Property Payment_Type() As String
        Get
            Return _PaymentType
        End Get
        Set(ByVal value As String)
            _PaymentType = value
        End Set
    End Property
    Public Property CollectionDate() As DateTime
        Get
            Return _CollectionDate
        End Get
        Set(ByVal value As DateTime)
            _CollectionDate = value
        End Set
    End Property
    Public Property Reference() As String
        Get
            Return _Reference
        End Get
        Set(ByVal value As String)
            _Reference = value
        End Set
    End Property
    Public Property CollectionAmount() As Decimal
        Get
            Return _CollectionAmount
        End Get
        Set(ByVal value As Decimal)
            _CollectionAmount = value
        End Set
    End Property
    Public Property Type() As String
        Get
            Return _Type
        End Get
        Set(ByVal value As String)
            _Type = value
        End Set
    End Property
    Public Property StockBalance() As Decimal
        Get
            Return _StockBalance
        End Get
        Set(ByVal value As Decimal)
            _StockBalance = value
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

    Public Property FromDate() As DateTime
        Get
            Return _FromDate
        End Get
        Set(ByVal value As DateTime)
            _FromDate = value
        End Set
    End Property
    Public Property ToDate() As DateTime
        Get
            Return _ToDate
        End Get
        Set(ByVal value As DateTime)
            _ToDate = Value
        End Set
    End Property
#End Region

#Region "Insert Purchase"
    Public Sub InsertPurchase(ByVal db As Database, ByVal transaction As DbTransaction)

        Try

            Dim DBC As DbCommand = db.GetStoredProcCommand(PURCHASES_INSERT)
            db.AddInParameter(DBC, "@PurchaseID", DbType.Int64, Me.PurchaseID)
            db.AddInParameter(DBC, "@SupplierID", DbType.Int64, Me.SupplierID)
            db.AddInParameter(DBC, "@ReferenceNo", DbType.String, Me.RefBillNo)
            db.AddInParameter(DBC, "@PurchaseDate", DbType.DateTime, Me.PurchaseDate)
            db.AddInParameter(DBC, "@Total", DbType.Decimal, Me.Total)
            db.AddInParameter(DBC, "@Percentage", DbType.Decimal, Me.TaxPercent)
            db.AddInParameter(DBC, "@Tax", DbType.Decimal, Me.TaxAmount)
            db.AddInParameter(DBC, "@Discount", DbType.Decimal, Me.Discount)
            db.AddInParameter(DBC, "@GrandTotal", DbType.Decimal, Me.GrandTotal)

            db.AddInParameter(DBC, "@Note", DbType.String, Me.Note)
            db.AddInParameter(DBC, "@Paid", DbType.Boolean, Me.Paid)
            db.AddOutParameter(DBC, "@CurrentPurchaseID", DbType.Int64, 0)
            db.AddInParameter(DBC, "@CreatedBy", DbType.Int64, Me.CreatedBy)
            db.AddInParameter(DBC, "@UpdatedBy", DbType.Int64, Me.UpdatedBy)

            db.ExecuteNonQuery(DBC, transaction)

            Me.CurrentPurchaseID = Convert.ToInt64(db.GetParameterValue(DBC, "@CurrentPurchaseID").ToString())
        Catch ex As Exception

            Throw
        End Try

    End Sub
#End Region

#Region "Insert Purchase Description"
    Public Sub InsertPurchaseDescription(ByVal db As Database, ByVal transaction As DbTransaction)

        Try

            Dim DBC As DbCommand = db.GetStoredProcCommand(PURCHASESDESCRIPTION_INSERT)
            db.AddInParameter(DBC, "@PurchaseID", DbType.Int64, Me.PurchaseID)
            db.AddInParameter(DBC, "@StockID", DbType.Int64, Me.StockID)
            db.AddInParameter(DBC, "@Unit_Price", DbType.Decimal, Me.Unit_Price)
            db.AddInParameter(DBC, "@Quantity", DbType.Decimal, Me.Quantity)
            db.AddInParameter(DBC, "@Discount", DbType.Decimal, Me.PDiscount)
            db.AddInParameter(DBC, "@Value", DbType.Decimal, Me.Value)
            db.ExecuteNonQuery(DBC, transaction)

        Catch ex As Exception
            Throw

        End Try

    End Sub
#End Region

#Region "Get Purchas By PurchaseID"
    Public Sub GetPurchasesByID()
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(PURCHASES_GETBYID)
            DB.AddInParameter(DBC, "@PurchaseID", DbType.Int64, Me.PurchaseID)

            Using DR As IDataReader = DB.ExecuteReader(DBC)


                With DR
                    Do While .Read

                        Me.SupplierID = Convert.ToInt64(IIf(Not IsDBNull(.Item("SupplierID")), Trim(.Item("SupplierID").ToString), 0))
                        Me.PurchaseNo = Convert.ToInt64(IIf(Not IsDBNull(.Item("PurchaseNo")), Trim(.Item("PurchaseNo").ToString), 0))
                        Me.RefBillNo = IIf(Not IsDBNull(.Item("ReferenceNo")), Trim(.Item("ReferenceNo").ToString), String.Empty)
                        Me.PurchaseDate = IIf(Not IsDBNull(.Item("PurchaseDate")), DateTime.Parse(Trim(.Item("PurchaseDate").ToString)), String.Empty)
                        Me.Total = Convert.ToDecimal(IIf(Not IsDBNull(.Item("Total")), .Item("Total").ToString, 0))
                        Me.TaxPercent = Convert.ToDecimal(IIf(Not IsDBNull(.Item("Percentage")), .Item("Percentage").ToString, 0))
                        Me.TaxAmount = Convert.ToDecimal(IIf(Not IsDBNull(.Item("Tax")), .Item("Tax").ToString, String.Empty))
                        Me.Discount = Convert.ToDecimal(IIf(Not IsDBNull(.Item("Discount")), .Item("Discount").ToString, 0))
                        Me.GrandTotal = Convert.ToDecimal(IIf(Not IsDBNull(.Item("GrandTotal")), .Item("GrandTotal").ToString, 0))

                        Me.Note = IIf(Not IsDBNull(.Item("Note")), Trim(.Item("Note").ToString), String.Empty)
                        Me.Paid = IIf(Not IsDBNull(.Item("Paid")), Boolean.Parse(Trim(.Item("Paid").ToString)), String.Empty)
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
        End Try
    End Sub
#End Region

#Region "Get PurchaseDescription By PurchaseID"
    Public Function GetPurchaseDescriptionByID() As DataSet
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(PURCHASESDESCRIPTION_GETBYID)
            DB.AddInParameter(DBC, "@PurchaseID", DbType.Int64, Me.PurchaseID)
            Return DB.ExecuteDataSet(DBC)

        Catch ex As Exception
            Return Nothing
        End Try
    End Function
#End Region

#Region "GetAll Purchases"
    Public Function GetPurchasesByDates() As DataSet
        Try

            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(PURCHASES_GETBYDATES)
            DB.AddInParameter(DBC, "@FromDate", DbType.DateTime, Me.FromDate)
            DB.AddInParameter(DBC, "@ToDate", DbType.DateTime, Me.ToDate)

            Return DB.ExecuteDataSet(DBC)

        Catch ex As Exception
            Return Nothing
        End Try

    End Function
#End Region

#Region "Purchases Description Delete"
    Public Sub PurchasesDescriptionDelete(ByVal db As Database, ByVal transaction As DbTransaction)

        Try
            Dim DBC As DbCommand = db.GetStoredProcCommand(PURCHASESDESCRIPTION_DELETE)
            db.AddInParameter(DBC, "@PurchaseID", DbType.Int64, Me.PurchaseID)
            db.ExecuteNonQuery(DBC, transaction)

        Catch ex As Exception
            Throw

        End Try

    End Sub
#End Region

#Region "Remove from Stock from Purchase"
    Public Sub RemoveFromStock(ByVal db As Database, ByVal transaction As DbTransaction)

        Try

            Dim DBC As DbCommand = db.GetStoredProcCommand(STOCK_REMOVEFROMPURCHASE)
            db.AddInParameter(DBC, "@PurchaseID", DbType.Int64, Me.PurchaseID)
            db.ExecuteNonQuery(DBC, transaction)

        Catch ex As Exception
            Throw

        End Try

    End Sub
#End Region

#Region "Update Stock from Purchase"
    Public Sub UpdateStock(ByVal db As Database, ByVal transaction As DbTransaction)

        Try

            Dim DBC As DbCommand = db.GetStoredProcCommand(STOCK_UPDATEBYPURCHASE)

            db.AddInParameter(DBC, "@StockID", DbType.Int64, Me.StockID)
            db.AddInParameter(DBC, "@StockBalance", DbType.Decimal, Me.StockBalance)
            db.ExecuteNonQuery(DBC, transaction)

        Catch ex As Exception
            Throw

        End Try

    End Sub
#End Region

#Region "Get All Unsettled Purchase Bills"
    Public Function GetAllPurchaseUnsettledBills() As DataSet
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(PURCHASES_UNSETTLED_BILLS)
            Return DB.ExecuteDataSet(DBC)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
#End Region

#Region "Get All Settled Purchase Bills"
    Public Function GetAllPurchaseSettledBills() As DataSet
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(PURCHASES_SETTLED_BILLS)
            Return DB.ExecuteDataSet(DBC)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
#End Region

#Region "Pruchase Set as paid"
    Public Sub PurchaseSetAsPaid()

        Try

            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(PURCHASES_SETASPAID)
            
            DB.AddInParameter(DBC, "@PurchaseID", DbType.Int64, Me.PurchaseID)
            DB.AddInParameter(DBC, "@UpdatedBy", DbType.Int64, Me.UpdatedBy)
            DB.ExecuteNonQuery(DBC)

        Catch ex As Exception
            Throw

        End Try

    End Sub
#End Region

End Class
