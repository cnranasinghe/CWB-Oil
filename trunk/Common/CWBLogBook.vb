Imports Microsoft.Practices.EnterpriseLibrary.Common
Imports Microsoft.Practices.EnterpriseLibrary.Data
Imports System.Data
Imports System.Data.Common
Imports CWBCommon.CWBConstants
Public Class CWBLogBook

#Region "Variables"

    Private _UserName As String
    Private _Password As String
    Private _AttempDate As DateTime
    Private _IPAddress As String
    Private _ComputerName As String

    Private _FromDate As DateTime
    Private _ToDate As DateTime

#End Region

#Region "Properties"
    Public Property UserName() As String
        Get
            Return _UserName
        End Get
        Set(ByVal value As String)
            _UserName = value
        End Set
    End Property
    Public Property Password() As String
        Get
            Return _Password
        End Get
        Set(ByVal value As String)
            _Password = Value
        End Set
    End Property
    Public Property AttempDate() As DateTime
        Get
            Return _AttempDate
        End Get
        Set(ByVal value As DateTime)
            _AttempDate = value
        End Set
    End Property
    Public Property IPAddress() As String
        Get
            Return _IPAddress
        End Get
        Set(ByVal value As String)
            _IPAddress = value
        End Set
    End Property
    Public Property ComputerName() As String
        Get
            Return _ComputerName
        End Get
        Set(ByVal value As String)
            _ComputerName = value
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

#Region "Logbook Insert"
    Public Sub InsertLogbook()

        Try

            Dim db As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = db.GetStoredProcCommand(LOGBOOK_INSERT)

            db.AddInParameter(DBC, "@UserName", DbType.String, Me.UserName)
            db.AddInParameter(DBC, "@Password", DbType.String, Me.Password)
            db.AddInParameter(DBC, "@IPAddress", DbType.String, Me.IPAddress)
            db.AddInParameter(DBC, "@ComputerName", DbType.String, Me.ComputerName)

            db.ExecuteNonQuery(DBC)

      
          
        Catch ex As Exception
            Throw

        End Try

    End Sub
#End Region

#Region "Get Logbook By Dates"
    Public Function GetLogBookByDates() As DataSet
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(LOGBOOK_GETBYDATES)
            DB.AddInParameter(DBC, "@FromDate", DbType.DateTime, Me.FromDate)
            DB.AddInParameter(DBC, "@ToDate", DbType.DateTime, Me.ToDate)

            Return DB.ExecuteDataSet(DBC)
            DBC.Dispose()
        Catch ex As Exception

            Return Nothing
            Throw
        End Try




    End Function
#End Region

End Class
