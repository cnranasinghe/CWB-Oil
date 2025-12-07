Imports Microsoft.Practices.EnterpriseLibrary.Common
Imports Microsoft.Practices.EnterpriseLibrary.Data
Imports System.Data
Imports System.Data.Common
Imports CWBCommon.CWBConstants
Public Class CWBModels

#Region "Vairables"
    Private _ModelID As Int64
    Private _Description As String
#End Region

#Region "Property"
    Public Property ModelID() As Int64
        Get
            Return _ModelID
        End Get
        Set(ByVal value As Int64)
            _ModelID = value
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

    
#End Region

#Region "Insert Models"


    Public Sub InsertModels()
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(MODELS_INSERT)
            DB.AddInParameter(DBC, "@Description", DbType.String, Me.Description)
            DB.ExecuteNonQuery(DBC)
        Catch ex As Exception
            Throw

        End Try

    End Sub

#End Region

#Region "Delete Models"
    Public Sub DeleteModels()
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(MODELS_DELETE)
            DB.AddInParameter(DBC, "@Description", DbType.String, Me.Description)
            DB.ExecuteNonQuery(DBC)
        Catch ex As Exception
            Throw

        End Try

    End Sub
#End Region

#Region "Get All Models"
    Public Function GetAllModels() As DataSet
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(MODELS_GETALL)
            Return DB.ExecuteDataSet(DBC)
            DBC.Dispose()
        Catch ex As Exception

            Return Nothing
            Throw

        End Try




    End Function
#End Region

End Class
