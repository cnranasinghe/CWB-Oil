Imports Microsoft.Practices.EnterpriseLibrary.Common
Imports Microsoft.Practices.EnterpriseLibrary.Data
Imports System.Data
Imports System.Data.Common
Imports CWBCommon.CWBConstants

Public Class CWBCustomers

#Region "Variables"
    Private _CustomerID As Int64
    Private _CustomerNo As String
    Private _CustomerName As String
    Private _AddressLine1 As String
    Private _AddressLine2 As String
    Private _AddressLine3 As String
    Private _Telephone As String
    Private _Fax As String
    Private _Email As String
    Private _WebURL As String
    Private _CreatedBy As Int64
    Private _CreatedDate As DateTime
    Private _UpdatedBy As Int64
    Private _UpdatedDate As DateTime
#End Region

#Region "Properties"
    Public Property CustomerID() As Int64
        Get
            Return _CustomerID
        End Get
        Set(ByVal value As Int64)
            _CustomerID = value
        End Set
    End Property
    Public Property CustomerNo() As String
        Get
            Return _CustomerNo
        End Get
        Set(ByVal value As String)
            _CustomerNo = value
        End Set
    End Property
    Public Property CustomerName() As String
        Get
            Return _CustomerName
        End Get
        Set(ByVal value As String)
            _CustomerName = value
        End Set
    End Property
    Public Property AddressLine1() As String
        Get
            Return _AddressLine1
        End Get
        Set(ByVal value As String)
            _AddressLine1 = value
        End Set
    End Property
    Public Property AddressLine2() As String
        Get
            Return _AddressLine2
        End Get
        Set(ByVal value As String)
            _AddressLine2 = value
        End Set
    End Property
    Public Property AddressLine3() As String
        Get
            Return _AddressLine3
        End Get
        Set(ByVal value As String)
            _AddressLine3 = value
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

#Region "Insert Customers"
    Public Sub InsertCustomers()
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(CUSTOMERS_INSERT)
            DB.AddInParameter(DBC, "@CustomerID", DbType.Int64, Me.CustomerID)
            DB.AddInParameter(DBC, "@CustomerNo", DbType.String, Me.CustomerNo)
            DB.AddInParameter(DBC, "@CustomerName", DbType.String, Me.CustomerName)
            DB.AddInParameter(DBC, "@AddressLine1", DbType.String, Me.AddressLine1)
            DB.AddInParameter(DBC, "@AddressLine2", DbType.String, Me.AddressLine2)
            DB.AddInParameter(DBC, "@AddressLine3", DbType.String, Me.AddressLine3)
            DB.AddInParameter(DBC, "@Telephone", DbType.String, Me.Telephone)
            DB.AddInParameter(DBC, "@Fax", DbType.String, Me.Fax)
            DB.AddInParameter(DBC, "@Email", DbType.String, Me.Email)
            DB.AddInParameter(DBC, "@WebURL", DbType.String, Me.WebURL)
            DB.AddInParameter(DBC, "@CreatedBy", DbType.Int64, Me.CreatedBy)
            DB.AddInParameter(DBC, "@UpdatedBy", DbType.Int64, Me.UpdatedBy)
            DB.ExecuteNonQuery(DBC)
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#Region "Delete Customers"
    Public Sub DeleteCustomer()
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(CUSTOMERS_DELETE)
            DB.AddInParameter(DBC, "@CustomerID", DbType.Int64, Me.CustomerID)
            DB.ExecuteNonQuery(DBC)
        Catch ex As Exception
            Throw
        End Try

    End Sub
#End Region

#Region "Get Customer By CustomerID"
    Public Sub GetCustomerByID()
        Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
        Dim DBC As DbCommand = DB.GetStoredProcCommand(CUSTOMERS_GETBYID)
        Try

            DB.AddInParameter(DBC, "@CustomerID", DbType.Int64, Me.CustomerID)
            Using DR As IDataReader = DB.ExecuteReader(DBC)


                With DR
                    Do While .Read
                        Me.CustomerID = Convert.ToInt64(IIf(Not IsDBNull(.Item("CustomerID")), Trim(.Item("CustomerID").ToString), 0))
                        Me.CustomerNo = IIf(Not IsDBNull(.Item("CustomerNo")), Trim(.Item("CustomerNo").ToString), String.Empty)
                        Me.CustomerName = IIf(Not IsDBNull(.Item("CustomerName")), Trim(.Item("CustomerName").ToString), String.Empty)
                        Me.AddressLine1 = IIf(Not IsDBNull(.Item("AddressLine1")), Trim(.Item("AddressLine1").ToString), String.Empty)
                        Me.AddressLine2 = IIf(Not IsDBNull(.Item("AddressLine2")), Trim(.Item("AddressLine2").ToString), String.Empty)
                        Me.AddressLine3 = IIf(Not IsDBNull(.Item("AddressLine3")), Trim(.Item("AddressLine3").ToString), String.Empty)
                        Me.Telephone = IIf(Not IsDBNull(.Item("Telephone")), Trim(.Item("Telephone").ToString), String.Empty)
                        Me.Fax = IIf(Not IsDBNull(.Item("Fax")), Trim(.Item("Fax").ToString), String.Empty)
                        Me.Email = IIf(Not IsDBNull(.Item("Email")), Trim(.Item("Email").ToString), String.Empty)
                        Me.WebURL = IIf(Not IsDBNull(.Item("WebURL")), Trim(.Item("WebURL").ToString), String.Empty)
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

#Region "Get All Customers"
    Public Function GetAllCustomers() As DataSet
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(CUSTOMERS_GETALL)
            Return DB.ExecuteDataSet(DBC)
            DBC.Dispose()
        Catch ex As Exception

            Return Nothing
            Throw
        End Try




    End Function
#End Region

End Class
