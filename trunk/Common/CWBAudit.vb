Imports Microsoft.Practices.EnterpriseLibrary.Common
Imports Microsoft.Practices.EnterpriseLibrary.Data
Imports System.Data
Imports System.Data.Common
Imports CWBCommon.CWBConstants

Public Class CWBAudit
#Region "Variables"



    Private _FromDate As DateTime
    Private _ToDate As DateTime

#End Region

#Region "Properties"


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
#End Region

#Region "Get Audit by SalesID"
    Public Function GetSalesAuditBySalesID(salesno As Int64) As DataSet
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(SALES_GETYBSALESID)
            'DB.AddInParameter(DBC, "@FromDate", DbType.DateTime, Me.FromDate)
            'DB.AddInParameter(DBC, "@ToDate", DbType.DateTime, Me.ToDate)
            DB.AddInParameter(DBC, "@SalesNo", DbType.Int64, salesno)

            Return DB.ExecuteDataSet(DBC)
            DBC.Dispose()
        Catch ex As Exception

            Return Nothing
            Throw
        End Try

    End Function
#End Region

End Class
