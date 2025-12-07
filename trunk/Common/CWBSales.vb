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
    Private _MessageID As Int64

    Private _SalesNo As Int64

    Private _ReferenceNo As String
    Private _SalesDate As DateTime

    Private _FromDate As DateTime
    Private _ToDate As DateTime

    Private _Total As Decimal

    Private _ModelID As Int64
    Private _CurrenMeterReading As Decimal
    Private _NextMeterReading As Decimal

    Private _Service1 As Int64
    Private _Service2 As Int64




    Private _TaxPercent As Decimal
    Private _TaxAmount As Decimal
    Private _Discount As Decimal
    Private _GrandTotal As Decimal
    Private _Note As String
    Private _Paid As Boolean

    Private _Ch1 As Boolean
    Private _Ch2 As Boolean
    Private _Ch3 As Boolean
    Private _Ch4 As Boolean

    Private _UnitPrice As Decimal
    Private _PurchasingPrice As Decimal
    Private _Quantity As Decimal
    Private _PDiscount As Decimal
    Private _Value As Decimal
    Private _StockBalance As Decimal
    Private _CreatedBy As Int64
    Private _CreatedDate As DateTime
    Private _UpdatedBy As Int64
    Private _UpdatedDate As DateTime

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
    Public Property Service1() As Int64
        Get
            Return _Service1
        End Get
        Set(ByVal value As Int64)
            _Service1 = value
        End Set
    End Property
    Public Property Service2() As Int64
        Get
            Return _Service2
        End Get
        Set(ByVal value As Int64)
            _Service2 = value
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
            _ToDate = value
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

    Public Property ModelID() As Int64
        Get
            Return _ModelID
        End Get
        Set(ByVal value As Int64)
            _ModelID = value
        End Set
    End Property

    Public Property MessageID() As Int64
        Get
            Return _MessageID
        End Get
        Set(ByVal value As Int64)
            _MessageID = value
        End Set
    End Property



    Public Property SalesNo() As Int64
        Get
            Return _SalesNo
        End Get
        Set(ByVal value As Int64)
            _SalesNo = value
        End Set
    End Property
    Public Property CurrenMeterReading() As Decimal
        Get
            Return _CurrenMeterReading
        End Get
        Set(ByVal value As Decimal)
            _CurrenMeterReading = value
        End Set
    End Property

    Public Property NextMeterReading() As Decimal
        Get
            Return _NextMeterReading
        End Get
        Set(ByVal value As Decimal)
            _NextMeterReading = value
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
    Public Property Ch1() As Boolean
        Get
            Return _Ch1
        End Get
        Set(ByVal value As Boolean)
            _Ch1 = value
        End Set
    End Property
    Public Property Ch2() As Boolean
        Get
            Return _Ch2
        End Get
        Set(ByVal value As Boolean)
            _Ch2 = value
        End Set
    End Property
    Public Property Ch3() As Boolean
        Get
            Return _Ch3
        End Get
        Set(ByVal value As Boolean)
            _Ch3 = value
        End Set
    End Property
    Public Property Ch4() As Boolean
        Get
            Return _Ch4
        End Get
        Set(ByVal value As Boolean)
            _Ch4 = value
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

#Region "Insert Sales"
    Public Sub InsertSales(ByVal db As Database, ByVal transaction As DbTransaction)

        Try

            Dim DBC As DbCommand = db.GetStoredProcCommand(SALES_INSERT)
            db.AddInParameter(DBC, "@SalesID", DbType.Int64, Me.SalesID)
            db.AddInParameter(DBC, "@CustomerID", DbType.Int64, Me.CustomerID)
            db.AddInParameter(DBC, "@ReferenceNo", DbType.String, Me.ReferenceNo)
            db.AddInParameter(DBC, "@SalesDate", DbType.DateTime, Me.SalesDate)

            db.AddInParameter(DBC, "@ModelID", DbType.Int64, Me.ModelID)
            db.AddInParameter(DBC, "@CurrenMeterReading", DbType.Decimal, Me.CurrenMeterReading)
            db.AddInParameter(DBC, "@NextMeterReading", DbType.Decimal, Me.NextMeterReading)


            db.AddInParameter(DBC, "@Total", DbType.Decimal, Me.Total)
            db.AddInParameter(DBC, "@Percentage", DbType.Decimal, Me.TaxPercent)
            db.AddInParameter(DBC, "@Tax", DbType.Decimal, Me.TaxAmount)
            db.AddInParameter(DBC, "@Discount", DbType.Decimal, Me.Discount)
            db.AddInParameter(DBC, "@GrandTotal", DbType.Decimal, Me.GrandTotal)

            db.AddInParameter(DBC, "@Service1", DbType.Int64, Me.Service1)
            db.AddInParameter(DBC, "@Service2", DbType.Int64, Me.Service2)


            db.AddInParameter(DBC, "@Note", DbType.String, Me.Note)
            db.AddInParameter(DBC, "@Paid", DbType.Boolean, Me.Paid)
            db.AddInParameter(DBC, "@Ch1", DbType.Boolean, Me.Ch1)
            db.AddInParameter(DBC, "@Ch2", DbType.Boolean, Me.Ch2)
            db.AddInParameter(DBC, "@Ch3", DbType.Boolean, Me.Ch3)
            db.AddInParameter(DBC, "@Ch4", DbType.Boolean, Me.Ch4)
            db.AddInParameter(DBC, "@MessageID", DbType.Int64, Me.MessageID)

            db.AddOutParameter(DBC, "@CurrentSalesID", DbType.Int64, 0)
            db.AddInParameter(DBC, "@CreatedBy", DbType.Int64, Me.CreatedBy)
            db.AddInParameter(DBC, "@UpdatedBy", DbType.Int64, Me.UpdatedBy)
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

            Throw
        End Try

    End Sub
#End Region

#Region "Get Sales By SalesID"
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
                        Me.ModelID = Convert.ToInt64(IIf(Not IsDBNull(.Item("ModelID")), .Item("ModelID").ToString, 0))
                        Me.SalesNo = Convert.ToInt64(IIf(Not IsDBNull(.Item("SalesNo")), .Item("SalesNo").ToString, 0))
                        Me.CurrenMeterReading = Convert.ToDecimal(IIf(Not IsDBNull(.Item("CurrenMeterReading")), .Item("CurrenMeterReading").ToString, 0))
                        Me.NextMeterReading = Convert.ToDecimal(IIf(Not IsDBNull(.Item("NextMeterReading")), .Item("NextMeterReading").ToString, 0))


                        Me.TaxPercent = Convert.ToDecimal(IIf(Not IsDBNull(.Item("Percentage")), .Item("Percentage").ToString, 0))
                        Me.TaxAmount = Convert.ToDecimal(IIf(Not IsDBNull(.Item("Tax")), .Item("Tax").ToString, String.Empty))
                        Me.Discount = Convert.ToDecimal(IIf(Not IsDBNull(.Item("Discount")), .Item("Discount").ToString, 0))
                        Me.GrandTotal = Convert.ToDecimal(IIf(Not IsDBNull(.Item("GrandTotal")), .Item("GrandTotal").ToString, 0))

                        Me.Service1 = Convert.ToInt64(IIf(Not IsDBNull(.Item("Service1")), .Item("Service1").ToString, 0))
                        Me.Service2 = Convert.ToInt64(IIf(Not IsDBNull(.Item("Service2")), .Item("Service2").ToString, 0))


                        Me.Note = IIf(Not IsDBNull(.Item("Note")), Trim(.Item("Note").ToString), String.Empty)

                        Me.MessageID = Convert.ToInt64(IIf(Not IsDBNull(.Item("MessageID")), .Item("MessageID").ToString, 0))

                        Me.Ch1 = IIf(Not IsDBNull(.Item("Ch1")), Boolean.Parse(.Item("Ch1")), False)
                        Me.Ch2 = IIf(Not IsDBNull(.Item("Ch2")), Boolean.Parse(.Item("Ch2")), False)
                        Me.Ch3 = IIf(Not IsDBNull(.Item("Ch3")), Boolean.Parse(.Item("Ch3")), False)
                        Me.Ch4 = IIf(Not IsDBNull(.Item("Ch4")), Boolean.Parse(.Item("Ch4")), False)

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

#Region "Get SalesDescription By SalesID"
    Public Function GetSalesDescriptionByID() As DataSet
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(SALESDESCRIPTION_GETBYID)
            DB.AddInParameter(DBC, "@SalesID", DbType.Int64, Me.SalesID)
            Return DB.ExecuteDataSet(DBC)

        Catch ex As Exception
            Return Nothing
            Throw
        End Try
    End Function
#End Region

#Region "GetAll Sales"
    Public Function GetSalesByDates() As DataSet
        Try

            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(SALES_GETBYDATES)
            DB.AddInParameter(DBC, "@FromDate", DbType.DateTime, Me.FromDate)
            DB.AddInParameter(DBC, "@ToDate", DbType.DateTime, Me.ToDate)
            Return DB.ExecuteDataSet(DBC)

        Catch ex As Exception
            Return Nothing
            Throw
        End Try

    End Function
#End Region

#Region "Get Sales By ReferenceNo"
    Public Function GetSalesByReferenceNo() As DataSet
        Try

            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(SALES_GETBYREFERENCENO)
            DB.AddInParameter(DBC, "@ReferenceNo", DbType.String, Me.ReferenceNo)
            Return DB.ExecuteDataSet(DBC)

        Catch ex As Exception
            Return Nothing
            Throw
        End Try

    End Function
#End Region

#Region "Get Sales By ReferenceNo, Customer, From Date , To Date"
    Public Function GetSalesByReferenceNoCustomerAndDates() As DataSet
        Try

            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(SALES_GETBYREFERENCENOCUSTOMERANDDATES)
            DB.AddInParameter(DBC, "@ReferenceNo", DbType.String, Me.ReferenceNo)
            DB.AddInParameter(DBC, "@CustomerID", DbType.Int64, Me.CustomerID)
            DB.AddInParameter(DBC, "@FromDate", DbType.DateTime, Me.FromDate)
            DB.AddInParameter(DBC, "@ToDate", DbType.DateTime, Me.ToDate)

            Return DB.ExecuteDataSet(DBC)

        Catch ex As Exception
            Return Nothing
            Throw
        End Try

    End Function
#End Region

#Region "Sales Description Delete"
    Public Sub SalesDescriptionDelete(ByVal db As Database, ByVal transaction As DbTransaction)

        Try

            Dim DBC As DbCommand = db.GetStoredProcCommand(SALESDESCRIPTION_DELETE)

            DB.AddInParameter(DBC, "@SalesID", DbType.Int64, Me.SalesID)
            db.ExecuteNonQuery(DBC, transaction)

        Catch ex As Exception
            Throw

        End Try

    End Sub
#End Region

#Region "Add to Stock from Sales"
    Public Sub AddToStock(ByVal db As Database, ByVal transaction As DbTransaction)

        Try

            Dim DBC As DbCommand = db.GetStoredProcCommand(STOCK_ADDFROMSALES)
            DB.AddInParameter(DBC, "@SalesID", DbType.Int64, Me.SalesID)
            db.ExecuteNonQuery(DBC, transaction)

        Catch ex As Exception
            Throw

        End Try

    End Sub
#End Region

#Region "Update Stock from Sales"
    Public Sub UpdateStock(ByVal db As Database, ByVal transaction As DbTransaction)

        Try

            Dim DBC As DbCommand = db.GetStoredProcCommand(STOCK_UPDATEBYSALES)

            DB.AddInParameter(DBC, "@StockID", DbType.Int64, Me.StockID)
            DB.AddInParameter(DBC, "@StockBalance", DbType.Decimal, Me.StockBalance)
            db.ExecuteNonQuery(DBC, transaction)

        Catch ex As Exception

            Throw
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
            Throw
        End Try
    End Function
#End Region

#Region "Get All Unsettled Sales Bills"
    Public Function GetAllSalesUnsettledBillsByCustomerIDAndDates() As DataSet
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(SALES_UNSETTLED_BILLS_GEYBYCUSTOMERIDANDDATES)
            DB.AddInParameter(DBC, "@CustomerID", DbType.Int64, Me.CustomerID)
            DB.AddInParameter(DBC, "@FromDate", DbType.DateTime, Me.FromDate)
            DB.AddInParameter(DBC, "@ToDate", DbType.DateTime, Me.ToDate)

            Return DB.ExecuteDataSet(DBC)
        Catch ex As Exception
            Return Nothing
            Throw
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
            Throw
        End Try
    End Function
#End Region

#Region "Sales Set as paid"
    Public Sub SalesSetAsPaid()

        Try

            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(SALES_SETASPAID)

            DB.AddInParameter(DBC, "@SalesID", DbType.Int64, Me.SalesID)
            DB.AddInParameter(DBC, "@UpdatedBy", DbType.Int64, Me.UpdatedBy)
            DB.ExecuteNonQuery(DBC)

        Catch ex As Exception
            Throw

        End Try

    End Sub
#End Region

#Region "Sales update discount"
    Public Sub SalesUpdateDiscount()

        Try

            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(SALES_UPDATEDISCOUNT)

            DB.AddInParameter(DBC, "@SalesID", DbType.Int64, Me.SalesID)
            DB.AddInParameter(DBC, "@Percentage", DbType.Decimal, Me.PDiscount)
            DB.AddInParameter(DBC, "@Discount", DbType.Decimal, Me.Discount)
            DB.AddInParameter(DBC, "@GrandTotal", DbType.Decimal, Me.GrandTotal)
            DB.AddInParameter(DBC, "@UpdatedBy", DbType.Int64, Me.UpdatedBy)
            DB.ExecuteNonQuery(DBC)

        Catch ex As Exception
            Throw

        End Try

    End Sub
#End Region




End Class
