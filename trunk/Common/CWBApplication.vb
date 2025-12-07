Imports Microsoft.Practices.EnterpriseLibrary.Common
Imports Microsoft.Practices.EnterpriseLibrary.Data
Imports System.Data
Imports System.Data.Common
Imports CWBCommon.CWBConstants
Public Class CWBApplication

#Region "Variables"
    Private _ServerDateTime As DateTime
#End Region

#Region "Properties"
    Public Property ServerDateTime() As DateTime
        Get
            Return _ServerDateTime
        End Get
        Set(ByVal value As DateTime)
            _ServerDateTime = value
        End Set
    End Property
#End Region

#Region "Get ServerDatetime"
    Public Function GetServerDateTime() As DateTime
        Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
        Dim DBC As DbCommand = DB.GetStoredProcCommand(APPLICATION_GETSERVERDATETIME)
        Try



            DB.AddOutParameter(DBC, "@ServerDateTime", DbType.DateTime, 0)
            DB.ExecuteNonQuery(DBC)

            ServerDateTime = Convert.ToDateTime(DB.GetParameterValue(DBC, "@ServerDateTime").ToString())

            Return ServerDateTime

        Catch ex As Exception
            Throw
            Return DateTime.MinValue

     

        End Try


    End Function
#End Region

End Class
