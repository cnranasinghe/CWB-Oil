Imports Microsoft.Practices.EnterpriseLibrary.Common
Imports Microsoft.Practices.EnterpriseLibrary.Data
Imports System.Data
Imports System.Data.Common
Imports CWBCommon.CWBConstants
Public Class CWBMessages
#Region "Vairables"
    Private _MessageID As Int64
    Private _MessageText As String
#End Region

#Region "Property"
    Public Property MessageID() As Int64
        Get
            Return _MessageID
        End Get
        Set(ByVal value As Int64)
            _MessageID = value
        End Set
    End Property
    Public Property MessageText() As String
        Get
            Return _MessageText
        End Get
        Set(ByVal value As String)
            _MessageText = value
        End Set
    End Property


#End Region

#Region "Insert Message"


    Public Sub InsertMessage()
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(MESSAGES_INSERT)
            DB.AddInParameter(DBC, "@Message", DbType.String, Me.MessageText)
            DB.ExecuteNonQuery(DBC)
        Catch ex As Exception
            Throw

        End Try

    End Sub

#End Region

#Region "Delete Message"
    Public Sub DeleteMessage()
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(MESSAGES_DELETE)
            DB.AddInParameter(DBC, "@Message", DbType.String, Me.MessageText)
            DB.ExecuteNonQuery(DBC)
        Catch ex As Exception
            Throw

        End Try

    End Sub
#End Region

#Region "Get All Message"
    Public Function GetAllMessage() As DataSet
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(MESSAGES_GETALL)
            Return DB.ExecuteDataSet(DBC)
            DBC.Dispose()
        Catch ex As Exception

            Return Nothing
            Throw

        End Try




    End Function
#End Region
End Class
