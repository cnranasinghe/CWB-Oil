Imports Microsoft.Practices.EnterpriseLibrary.Common
Imports Microsoft.Practices.EnterpriseLibrary.Data
Imports System.Data
Imports System.Data.Common
Imports CWBCommon.CWBConstants
Public Class CWBCompany

#Region "Variables"
    Private _CompanyName As String
    Private _Address1 As String
    Private _Address2 As String
    Private _Telephone As String
    Private _Fax As String
    Private _Email As String
    Private _WebURL As String
    Private _RegNo As String

#End Region

#Region "Properties"
    Public Property CompanyName() As String
        Get
            Return _CompanyName
        End Get
        Set(ByVal value As String)
            _CompanyName = value
        End Set
    End Property
    Public Property Address1() As String
        Get
            Return _Address1
        End Get
        Set(ByVal value As String)
            _Address1 = value
        End Set
    End Property
    Public Property Address2() As String
        Get
            Return _Address2
        End Get
        Set(ByVal value As String)
            _Address2 = value
        End Set
    End Property
    Public Property Telephone() As String
        Get
            Return _Telephone
        End Get
        Set(ByVal value As String)
            _Telephone = value
        End Set
    End Property
    Public Property Fax() As String
        Get
            Return _Fax
        End Get
        Set(ByVal value As String)
            _Fax = value
        End Set
    End Property
    Public Property Email() As String
        Get
            Return _Email
        End Get
        Set(ByVal value As String)
            _Email = value
        End Set
    End Property
    Public Property WebURL() As String
        Get
            Return _WebURL
        End Get
        Set(ByVal value As String)
            _WebURL = value
        End Set
    End Property
    Public Property RegNo() As String
        Get
            Return _RegNo
        End Get
        Set(ByVal value As String)
            _RegNo = Value
        End Set
    End Property


#End Region

#Region "Insert Company"
    Public Sub InsertCompany()
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(COMPANY_INSERT)
            DB.AddInParameter(DBC, "@CompanyName", DbType.String, Me.CompanyName)
            DB.AddInParameter(DBC, "@Address1", DbType.String, Me.Address1)
            DB.AddInParameter(DBC, "@Address2", DbType.String, Me.Address2)
            DB.AddInParameter(DBC, "@Telephone", DbType.String, Me.Telephone)
            DB.AddInParameter(DBC, "@Fax", DbType.String, Me.Fax)
            DB.AddInParameter(DBC, "@Email", DbType.String, Me.Email)
            DB.AddInParameter(DBC, "@WebURL", DbType.String, Me.WebURL)
            DB.AddInParameter(DBC, "@RegNo", DbType.String, Me.RegNo)
            DB.ExecuteNonQuery(DBC)
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#Region "Get Company"
    Public Sub GetCompany()
        Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
        Dim DBC As DbCommand = DB.GetStoredProcCommand(COMPANY_GETALL)
        Try


            Using DR As IDataReader = DB.ExecuteReader(DBC)


                With DR
                    Do While .Read
                        Me.CompanyName = IIf(Not IsDBNull(.Item("CompanyName")), Trim(.Item("CompanyName").ToString), String.Empty)
                        Me.Address1 = IIf(Not IsDBNull(.Item("Address1")), Trim(.Item("Address1").ToString), String.Empty)
                        Me.Address2 = IIf(Not IsDBNull(.Item("Address2")), Trim(.Item("Address2").ToString), String.Empty)
                        Me.Telephone = IIf(Not IsDBNull(.Item("Telephone")), Trim(.Item("Telephone").ToString), String.Empty)
                        Me.Fax = IIf(Not IsDBNull(.Item("Fax")), Trim(.Item("Fax").ToString), String.Empty)
                        Me.Email = IIf(Not IsDBNull(.Item("Email")), Trim(.Item("Email").ToString), String.Empty)
                        Me.WebURL = IIf(Not IsDBNull(.Item("WebURL")), Trim(.Item("WebURL").ToString), String.Empty)
                        Me.RegNo = IIf(Not IsDBNull(.Item("RegNo")), Trim(.Item("RegNo").ToString), String.Empty)




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

End Class
