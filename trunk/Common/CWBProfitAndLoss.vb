Imports Microsoft.Practices.EnterpriseLibrary.Common
Imports Microsoft.Practices.EnterpriseLibrary.Data
Imports System.Data
Imports System.Data.Common
Imports CWBCommon.CWBConstants

Public Class CWBProfitAndLoss

#Region "Variables"

    Private _From As Date
    Private _To As Date
    Private _StockValue As Decimal
    Private _TotalPurchases As Decimal
    Private _TotalSales As Decimal
    Private _TotalPurchaseCost As Decimal
    Private _TotalExpenses As Decimal
    Private _TotalCashtoPay As Decimal
    Private _TotalCashtoReceive As Decimal



#End Region

#Region "Properties"

    Public Property FromDate() As Date
        Get
            Return _From
        End Get
        Set(ByVal value As Date)
            If _From = value Then
                Return
            End If
            _From = value
        End Set

    End Property

    Public Property ToDate() As Date
        Get
            Return _To
        End Get
        Set(ByVal value As Date)
            If _To = value Then
                Return
            End If
            _To = value
        End Set
    End Property

    Public Property StockValue() As Decimal
        Get
            Return _StockValue
        End Get
        Set(ByVal value As Decimal)
            If _StockValue = value Then
                Return
            End If
            _StockValue = value
        End Set
    End Property

    Public Property TotalPurchases() As Decimal
        Get
            Return _TotalPurchases
        End Get
        Set(ByVal value As Decimal)
            If _TotalPurchases = value Then
                Return
            End If
            _TotalPurchases = value
        End Set
    End Property

    Public Property TotalSales() As Decimal
        Get
            Return _TotalSales
        End Get
        Set(ByVal value As Decimal)
            If _TotalSales = value Then
                Return
            End If
            _TotalSales = value
        End Set
    End Property

    Public Property TotalPurchaseCost() As Decimal
        Get
            Return _TotalPurchaseCost
        End Get
        Set(ByVal value As Decimal)
            If _TotalPurchaseCost = value Then
                Return
            End If
            _TotalPurchaseCost = value
        End Set
    End Property

    Public Property TotalExpenses() As Decimal
        Get
            Return _TotalExpenses
        End Get
        Set(ByVal value As Decimal)
            If _TotalExpenses = value Then
                Return
            End If
            _TotalExpenses = value
        End Set
    End Property

    Public Property TotalCashtoPay() As Decimal
        Get
            Return _TotalCashtoPay
        End Get
        Set(ByVal value As Decimal)
            If _TotalCashtoPay = value Then
                Return
            End If
            _TotalCashtoPay = value
        End Set
    End Property

    Public Property TotalCashtoReceive() As Decimal
        Get
            Return _TotalCashtoReceive
        End Get
        Set(ByVal value As Decimal)
            If _TotalCashtoReceive = value Then
                Return
            End If
            _TotalCashtoReceive = value
        End Set
    End Property
#End Region

#Region "Get P&L by Dates"
    Public Sub GetPLByDates()
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(PROFOTANDLOSS_GETBYDATES)

            DB.AddInParameter(DBC, "@FromDate", DbType.DateTime, Me.FromDate)
            DB.AddInParameter(DBC, "@ToDate", DbType.DateTime, Me.ToDate)

            DB.AddOutParameter(DBC, "@StockValue", DbType.Decimal, 0)
            DB.AddOutParameter(DBC, "@TotalPurchases", DbType.Decimal, 0)
            DB.AddOutParameter(DBC, "@TotalSales", DbType.Decimal, 0)
            DB.AddOutParameter(DBC, "@TotalPurchaseCost", DbType.Decimal, 0)
            DB.AddOutParameter(DBC, "@TotalExpenses", DbType.Decimal, 0)
            DB.AddOutParameter(DBC, "@TotalCashtoPay", DbType.Decimal, 0)
            DB.AddOutParameter(DBC, "@TotalCashtoReceive", DbType.Decimal, 0)

            DB.ExecuteNonQuery(DBC)
           
            Me.StockValue = Convert.ToInt64(IIf(Not IsDBNull(DB.GetParameterValue(DBC, "@StockValue")), DB.GetParameterValue(DBC, "@StockValue"), 0))
            Me.TotalPurchases = Convert.ToInt64(IIf(Not IsDBNull(DB.GetParameterValue(DBC, "@TotalPurchases")), DB.GetParameterValue(DBC, "@TotalPurchases"), 0))
            Me.TotalSales = Convert.ToInt64(IIf(Not IsDBNull(DB.GetParameterValue(DBC, "@TotalSales")), DB.GetParameterValue(DBC, "@TotalSales"), 0))
            Me.TotalPurchaseCost = Convert.ToInt64(IIf(Not IsDBNull(DB.GetParameterValue(DBC, "@TotalPurchaseCost")), DB.GetParameterValue(DBC, "@TotalPurchaseCost"), 0))
            Me.TotalExpenses = Convert.ToInt64(IIf(Not IsDBNull(DB.GetParameterValue(DBC, "@TotalExpenses")), DB.GetParameterValue(DBC, "@TotalExpenses"), 0))
            Me.TotalCashtoPay = Convert.ToInt64(IIf(Not IsDBNull(DB.GetParameterValue(DBC, "@TotalCashtoPay")), DB.GetParameterValue(DBC, "@TotalCashtoPay"), 0))
            Me.TotalCashtoReceive = Convert.ToInt64(IIf(Not IsDBNull(DB.GetParameterValue(DBC, "@TotalCashtoReceive")), DB.GetParameterValue(DBC, "@TotalCashtoReceive"), 0))

        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region


End Class
