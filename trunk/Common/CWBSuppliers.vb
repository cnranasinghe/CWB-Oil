Imports Microsoft.Practices.EnterpriseLibrary.Common
Imports Microsoft.Practices.EnterpriseLibrary.Data
Imports System.Data
Imports System.Data.Common
Imports CWBCommon.CWBConstants
Public Class CWBSuppliers

#Region "Variables"
    Private _SupplierID As Int64
    Private _SupplierNo As String
    Private _Salutation As String
    Private _SupplierName As String
    Private _AddressLine1 As String
    Private _AddressLine2 As String
    Private _AddressLine3 As String
    Private _Telephone As String
    Private _Estate As String
    Private _AccountNo As String
    Private _Bank As String
    Private _CreatedBy As Int64
    Private _CreatedDate As DateTime
    Private _UpdatedBy As Int64
    Private _UpdatedDate As DateTime
#End Region

#Region "Properties"
    Public Property SupplierID() As Int64
        Get
            Return _SupplierID
        End Get
        Set(ByVal value As Int64)
            _SupplierID = value
        End Set
    End Property
    Public Property SupplierNo() As String
        Get
            Return _SupplierNo
        End Get
        Set(ByVal value As String)
            _SupplierNo = value
        End Set
    End Property
    Public Property Salutation() As String
        Get
            Return _Salutation
        End Get
        Set(ByVal value As String)
            _Salutation = Value
        End Set
    End Property
    Public Property SupplierName() As String
        Get
            Return _SupplierName
        End Get
        Set(ByVal value As String)
            _SupplierName = value
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
    Public Property Estate() As String
        Get
            Return _Estate
        End Get
        Set(ByVal value As String)
            _Estate = value
        End Set
    End Property
    Public Property AccountNo() As String
        Get
            Return _AccountNo
        End Get
        Set(ByVal value As String)
            _AccountNo = value
        End Set
    End Property
    Public Property Bank() As String
        Get
            Return _Bank
        End Get
        Set(ByVal value As String)
            _Bank = Value
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

#Region "Insert Suppliers"
    Public Sub InsertSuppliers()
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(SUPPLIERS_INSERT)
            DB.AddInParameter(DBC, "@SupplierID", DbType.Int64, Me.SupplierID)
            DB.AddInParameter(DBC, "@SupplierNo", DbType.String, Me.SupplierNo)
            DB.AddInParameter(DBC, "@Salutation", DbType.String, Me.Salutation)
            DB.AddInParameter(DBC, "@SupplierName", DbType.String, Me.SupplierName)
            DB.AddInParameter(DBC, "@AddressLine1", DbType.String, Me.AddressLine1)
            DB.AddInParameter(DBC, "@AddressLine2", DbType.String, Me.AddressLine2)
            DB.AddInParameter(DBC, "@AddressLine3", DbType.String, Me.AddressLine3)
            DB.AddInParameter(DBC, "@Telephone", DbType.String, Me.Telephone)
            DB.AddInParameter(DBC, "@Estate", DbType.String, Me.Estate)
            DB.AddInParameter(DBC, "@AccountNo", DbType.String, Me.AccountNo)
            DB.AddInParameter(DBC, "@Bank", DbType.String, Me.Bank)
            DB.AddInParameter(DBC, "@CreatedBy", DbType.Int64, Me.CreatedBy)
            DB.AddInParameter(DBC, "@UpdatedBy", DbType.Int64, Me.UpdatedBy)
            DB.ExecuteNonQuery(DBC)
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#Region "Is Supplier Exits"
    Public Function IsSupplierExits() As Boolean
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(SUPPLIERS_ISSUPPLIEREXISTS)

            DB.AddInParameter(DBC, "@SupplierNo", DbType.Int64, SupplierNo)
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

#Region "Delete Supplier"
    Public Sub DeleteSupplier()
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(SUPPLIERS_DELETE)
            DB.AddInParameter(DBC, "@SupplierID", DbType.Int64, Me.SupplierID)
            DB.ExecuteNonQuery(DBC)
        Catch ex As Exception
            Throw
        End Try

    End Sub



#End Region

#Region "Get Supplier By SupplierID"
    Public Sub GetSupplierByID()
        Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
        Dim DBC As DbCommand = DB.GetStoredProcCommand(SUPPLIERS_GETBYID)
        Try

            DB.AddInParameter(DBC, "@SupplierID", DbType.Int64, Me.SupplierID)
            Using DR As IDataReader = DB.ExecuteReader(DBC)


                With DR
                    Do While .Read
                        Me.SupplierID = Convert.ToInt64(IIf(Not IsDBNull(.Item("SupplierID")), Trim(.Item("SupplierID").ToString), 0))
                        Me.SupplierNo = IIf(Not IsDBNull(.Item("SupplierNo")), Trim(.Item("SupplierNo").ToString), String.Empty)
                        Me.Salutation = IIf(Not IsDBNull(.Item("Salutation")), Trim(.Item("Salutation").ToString), String.Empty)
                        Me.SupplierName = IIf(Not IsDBNull(.Item("SupplierName")), Trim(.Item("SupplierName").ToString), String.Empty)
                        Me.AddressLine1 = IIf(Not IsDBNull(.Item("AddressLine1")), Trim(.Item("AddressLine1").ToString), String.Empty)
                        Me.AddressLine2 = IIf(Not IsDBNull(.Item("AddressLine2")), Trim(.Item("AddressLine2").ToString), String.Empty)
                        Me.AddressLine3 = IIf(Not IsDBNull(.Item("AddressLine3")), Trim(.Item("AddressLine3").ToString), String.Empty)
                        Me.Telephone = IIf(Not IsDBNull(.Item("Telephone")), Trim(.Item("Telephone").ToString), String.Empty)
                        Me.Estate = IIf(Not IsDBNull(.Item("Estate")), Trim(.Item("Estate").ToString), String.Empty)
                        Me.AccountNo = IIf(Not IsDBNull(.Item("AccountNo")), Trim(.Item("AccountNo").ToString), String.Empty)
                        Me.Bank = IIf(Not IsDBNull(.Item("Bank")), Trim(.Item("Bank").ToString), String.Empty)
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

#Region "Get Supplier By SupplierNo"
    Public Sub GetSupplierBySupplierNo()
        Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
        Dim DBC As DbCommand = DB.GetStoredProcCommand(SUPPLIERS_GETBYSUPPLIERNO)
        Try

            DB.AddInParameter(DBC, "@SupplierNo", DbType.Int64, Me.SupplierNo)
            Using DR As IDataReader = DB.ExecuteReader(DBC)


                With DR
                    Do While .Read
                        Me.SupplierID = Convert.ToInt64(IIf(Not IsDBNull(.Item("SupplierID")), .Item("SupplierID").ToString, 0))
                        Me.SupplierNo = IIf(Not IsDBNull(.Item("SupplierNo")), Trim(.Item("SupplierNo").ToString), String.Empty)
                        Me.Salutation = IIf(Not IsDBNull(.Item("Salutation")), Trim(.Item("Salutation").ToString), String.Empty)
                        Me.SupplierName = IIf(Not IsDBNull(.Item("SupplierName")), Trim(.Item("SupplierName").ToString), String.Empty)
                        Me.AddressLine1 = IIf(Not IsDBNull(.Item("AddressLine1")), Trim(.Item("AddressLine1").ToString), String.Empty)
                        Me.AddressLine2 = IIf(Not IsDBNull(.Item("AddressLine2")), Trim(.Item("AddressLine2").ToString), String.Empty)
                        Me.AddressLine3 = IIf(Not IsDBNull(.Item("AddressLine3")), Trim(.Item("AddressLine3").ToString), String.Empty)
                        Me.Telephone = IIf(Not IsDBNull(.Item("Telephone")), Trim(.Item("Telephone").ToString), String.Empty)
                        Me.Estate = IIf(Not IsDBNull(.Item("Estate")), Trim(.Item("Estate").ToString), String.Empty)
                        Me.AccountNo = IIf(Not IsDBNull(.Item("AccountNo")), Trim(.Item("AccountNo").ToString), String.Empty)
                        Me.Bank = IIf(Not IsDBNull(.Item("Bank")), Trim(.Item("Bank").ToString), String.Empty)
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

#Region "Get All Suppliers"
    Public Function GetAllSuppliers() As DataSet
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(SUPPLIERS_GETALL)
            Return DB.ExecuteDataSet(DBC)
            DBC.Dispose()
        Catch ex As Exception

            Return Nothing
            Throw
        End Try




    End Function
#End Region

End Class
