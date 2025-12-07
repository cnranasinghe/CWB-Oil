Imports Microsoft.Practices.EnterpriseLibrary.Common
Imports Microsoft.Practices.EnterpriseLibrary.Data
Imports System.Data
Imports System.Data.Common
Imports CWBCommon.CWBConstants

Public Class CWBUserLogins

#Region "Variables"
    Private _UserLoginID As Int64
    Private _EmployerID As Int64
    Private _UserType As String
    Private _UserName As String
    Private _Password As String
    Private _CreatedBy As Int64
    Private _CreatedDate As Date
    Private _UpdatedBy As Int64
    Private _UpdateDate As Date


#End Region

#Region "Properties"
    Public Property UserLoginID() As Int64
        Get
            Return _UserLoginID
        End Get
        Set(ByVal value As Int64)
            _UserLoginID = value
        End Set
    End Property
    Public Property EmployerID() As Int64
        Get
            Return _EmployerID
        End Get
        Set(ByVal value As Int64)
            _EmployerID = value
        End Set
    End Property
    Public Property UserType() As String
        Get
            Return _UserType
        End Get
        Set(ByVal value As String)
            _UserType = value
        End Set
    End Property
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
            _Password = value
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
    Public Property CreatedDate() As Date
        Get
            Return _CreatedDate
        End Get
        Set(ByVal value As Date)
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
    Public Property UpdateDate() As Date
        Get
            Return _UpdateDate
        End Get
        Set(ByVal value As Date)
            _UpdateDate = value
        End Set
    End Property


#End Region

#Region "Is User Exits"
    Public Function IsUserNameExits() As Boolean
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(USERLOGIN_ISUSERNAMEEXISTS)

            DB.AddInParameter(DBC, "@UserName", DbType.String, UserName)
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

#Region "Insert User"
    Public Sub Insert()
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(USERLOGIN_INSERT)
            DB.AddInParameter(DBC, "@UserLoginID", DbType.Int64, Me.UserLoginID)
            DB.AddInParameter(DBC, "@EmployerID", DbType.Int64, Me.EmployerID)
            DB.AddInParameter(DBC, "@UserType", DbType.String, Me.UserType)
            DB.AddInParameter(DBC, "@UserName", DbType.String, Me.UserName)
            DB.AddInParameter(DBC, "@Password", DbType.String, Me.Password)
            DB.AddInParameter(DBC, "@CreatedBy", DbType.Int64, Me.CreatedBy)
            DB.AddInParameter(DBC, "@UpdatedBy", DbType.Int64, Me.UpdatedBy)

            DB.ExecuteNonQuery(DBC)
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#Region "Delete USer"
    Public Sub Delete()
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(USERLOGIN_DELETE)
            DB.AddInParameter(DBC, "@UserLoginID", DbType.Int64, Me.UserLoginID)
            DB.ExecuteNonQuery(DBC)
        Catch ex As Exception
            Throw
        End Try

    End Sub



#End Region

#Region "Select User Login ID"
    Public Sub SelectRow()
        Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
        Dim DBC As DbCommand = DB.GetStoredProcCommand(USERLOGIN_GETBYID)
        Try

            DB.AddInParameter(DBC, "@UserLoginID", DbType.Int64, Me.UserLoginID)
            Using DR As IDataReader = DB.ExecuteReader(DBC)


                With DR
                    Do While .Read
                        Me.UserLoginID = Convert.ToInt64(IIf(Not IsDBNull(.Item("UserLoginID")), Trim(.Item("UserLoginID").ToString), 0))
                        Me.EmployerID = Convert.ToInt64(IIf(Not IsDBNull(.Item("EmployerID")), Trim(.Item("EmployerID").ToString), 0))
                        Me.UserType = IIf(Not IsDBNull(.Item("UserType")), Trim(.Item("UserType").ToString), String.Empty)
                        Me.UserName = IIf(Not IsDBNull(.Item("UserName")), Trim(.Item("UserName").ToString), String.Empty)
                        Me.Password = IIf(Not IsDBNull(.Item("Password")), Trim(.Item("Password").ToString), String.Empty)
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

#Region "Select user By UserName and Password"
    Public Sub SelectRowByUserNameAndPassword()
        Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
        Dim DBC As DbCommand = DB.GetStoredProcCommand(USERLOGIN_GETBYUSERNAMEANDPASSWORD)
        Try

            DB.AddInParameter(DBC, "@UserName", DbType.String, Me.UserName)
            DB.AddInParameter(DBC, "@Password", DbType.String, Me.Password)


            Using DR As IDataReader = DB.ExecuteReader(DBC)


                With DR
                    Do While .Read
                        Me.UserLoginID = Convert.ToInt64(IIf(Not IsDBNull(.Item("UserLoginID")), Trim(.Item("UserLoginID").ToString), 0))
                        Me.EmployerID = Convert.ToInt64(IIf(Not IsDBNull(.Item("EmployerID")), Trim(.Item("EmployerID").ToString), 0))
                        Me.UserType = IIf(Not IsDBNull(.Item("UserType")), Trim(.Item("UserType").ToString), String.Empty)
                        Me.UserName = IIf(Not IsDBNull(.Item("UserName")), Trim(.Item("UserName").ToString), String.Empty)
                        Me.Password = IIf(Not IsDBNull(.Item("Password")), Trim(.Item("Password").ToString), String.Empty)
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

#Region "Select All Users"
    Public Function SelectAll() As DataSet
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(USERLOGIN_GETALL)
            Return DB.ExecuteDataSet(DBC)
            DBC.Dispose()
        Catch ex As Exception

            Return Nothing
            Throw
        End Try




    End Function
#End Region


End Class
