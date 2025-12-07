Imports Microsoft.Practices.EnterpriseLibrary.Common
Imports Microsoft.Practices.EnterpriseLibrary.Data
Imports System.Data
Imports System.Data.Common
Imports CWBCommon.CWBConstants
Imports System.Drawing



Public Class CWBEmployers

#Region "Variables"
    Private _EmployerID As Int64
    Private _EmployerNo As Int64
    Private _EmployerName As String
    Private _EmployerImage As Image
    Private _AddressLine1 As String
    Private _AddressLine2 As String
    Private _AddressLine3 As String
    Private _DateOfBirth As DateTime
    Private _DateJoined As DateTime
    Private _Sex As String
    Private _NICNo As String
    Private _TelephoneNo As String
    Private _Designation As String
    Private _Department As String
    Private _EmergencyContactPerson As String
    Private _BasicSalary As Decimal
    Private _OTRate As Decimal
    Private _FixedAllowance As Decimal
    Private _OtherAllowance As Decimal
    Private _Deductions As Decimal
    Private _EPFNo As String
    Private _CreatedBy As Int64
    Private _CreatedDate As Date
    Private _UpdatedBy As Int64
    Private _UpdatedDate As Date
#End Region

#Region "Properties"
    Public Property EmployerID() As Int64
        Get
            Return _EmployerID
        End Get
        Set(ByVal value As Int64)
            _EmployerID = value
        End Set
    End Property
    Public Property EmployerNo() As Int64
        Get
            Return _EmployerNo
        End Get
        Set(ByVal value As Int64)
            _EmployerNo = value
        End Set
    End Property
    Public Property EmployerName() As String
        Get
            Return _EmployerName
        End Get
        Set(ByVal value As String)
            _EmployerName = value
        End Set
    End Property
    Public Property EmployerImage() As Image
        Get
            Return _EmployerImage
        End Get
        Set(ByVal value As Image)
            _EmployerImage = Value
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
    Public Property DateOfBirth() As DateTime
        Get
            Return _DateOfBirth
        End Get
        Set(ByVal value As DateTime)
            _DateOfBirth = value
        End Set
    End Property
    Public Property DateJoined() As DateTime
        Get
            Return _DateJoined
        End Get
        Set(ByVal value As DateTime)
            _DateJoined = value
        End Set
    End Property
    Public Property Sex() As String
        Get
            Return _Sex
        End Get
        Set(ByVal value As String)
            _Sex = value
        End Set
    End Property
    Public Property NICNo() As String
        Get
            Return _NICNo
        End Get
        Set(ByVal value As String)
            _NICNo = value
        End Set
    End Property
    Public Property TelephoneNo() As String
        Get
            Return _TelephoneNo
        End Get
        Set(ByVal value As String)
            _TelephoneNo = value
        End Set
    End Property
    Public Property Designation() As String
        Get
            Return _Designation
        End Get
        Set(ByVal value As String)
            _Designation = value
        End Set
    End Property
    Public Property Department() As String
        Get
            Return _Department
        End Get
        Set(ByVal value As String)
            _Department = Value
        End Set
    End Property
    Public Property EmergencyContactPerson() As String
        Get
            Return _EmergencyContactPerson
        End Get
        Set(ByVal value As String)
            _EmergencyContactPerson = value
        End Set
    End Property
    Public Property BasicSalary() As Decimal
        Get
            Return _BasicSalary
        End Get
        Set(ByVal value As Decimal)
            _BasicSalary = value
        End Set
    End Property
    Public Property OTRate() As Decimal
        Get
            Return _OTRate
        End Get
        Set(ByVal value As Decimal)
            _OTRate = value
        End Set
    End Property
    Public Property FixedAllowance() As Decimal
        Get
            Return _FixedAllowance
        End Get
        Set(ByVal value As Decimal)
            _FixedAllowance = value
        End Set
    End Property
    Public Property OtherAllowance() As Decimal
        Get
            Return _OtherAllowance
        End Get
        Set(ByVal value As Decimal)
            _OtherAllowance = value
        End Set
    End Property
    Public Property Deductions() As Decimal
        Get
            Return _Deductions
        End Get
        Set(ByVal value As Decimal)
            _Deductions = value
        End Set
    End Property
    Public Property EPFNo() As String
        Get
            Return _EPFNo
        End Get
        Set(ByVal value As String)
            _EPFNo = value
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
    Public Property UpdatedDate() As Date
        Get
            Return _UpdatedDate
        End Get
        Set(ByVal value As Date)
            _UpdatedDate = value
        End Set
    End Property
#End Region

#Region "Is Employer Exits"
    Public Function IsEmployerExits() As Boolean
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(EMPLOYERS_ISEMPLOYEREXISTS)

            DB.AddInParameter(DBC, "@EmployerNo", DbType.Int64, EmployerNo)
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

#Region "Getting Byte array from the image"
    Public Function GetByteArrayFromImage(ByVal img As Bitmap) As Byte()
        Try

            Dim ms As New System.IO.MemoryStream
            img.Save(ms, Imaging.ImageFormat.Bmp)
            Dim outBytes(CInt(ms.Length - 1)) As Byte
            ms.Seek(0, System.IO.SeekOrigin.Begin)
            ms.Read(outBytes, 0, CInt(ms.Length))
            Return outBytes
        Catch ex As Exception
            Return Nothing
            Throw

        End Try

    End Function
#End Region

#Region "Insert Employer"
    Public Sub Insert()
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(EMPLOYERS_INSERT)
            DB.AddInParameter(DBC, "@EmployerID", DbType.Int64, EmployerID)
            DB.AddInParameter(DBC, "@EmployerNo", DbType.String, EmployerNo)
            DB.AddInParameter(DBC, "@EmployerName", DbType.String, EmployerName)
            DB.AddInParameter(DBC, "@EmergencyContactPerson", DbType.String, EmergencyContactPerson)

            If EmployerImage Is Nothing Then
                DB.AddInParameter(DBC, "@EmployerImage", DbType.Binary, Nothing)
            Else
                DB.AddInParameter(DBC, "@EmployerImage", DbType.Binary, GetByteArrayFromImage(EmployerImage))

            End If

            DB.AddInParameter(DBC, "@AddressLine1", DbType.String, AddressLine1)
            DB.AddInParameter(DBC, "@AddressLine2", DbType.String, AddressLine2)
            DB.AddInParameter(DBC, "@AddressLine3", DbType.String, AddressLine3)

            If DateOfBirth <> Date.MinValue Then
                DB.AddInParameter(DBC, "@DateOfBirth", DbType.Date, DateOfBirth)
            Else
                DB.AddInParameter(DBC, "@DateOfBirth", DbType.Date, DBNull.Value)
            End If

            If DateJoined <> Date.MinValue Then
                DB.AddInParameter(DBC, "@DateJoined", DbType.Date, DateJoined)
            Else
                DB.AddInParameter(DBC, "@DateJoined", DbType.Date, DBNull.Value)

            End If

            DB.AddInParameter(DBC, "@Sex", DbType.String, Sex)
            DB.AddInParameter(DBC, "@NICNo", DbType.String, NICNo)
            DB.AddInParameter(DBC, "@TelephoneNo", DbType.String, TelephoneNo)
            DB.AddInParameter(DBC, "@Designation", DbType.String, Designation)
            DB.AddInParameter(DBC, "@Department", DbType.String, Department)
            DB.AddInParameter(DBC, "@BasicSalary", DbType.Decimal, BasicSalary)
            DB.AddInParameter(DBC, "@OTRate", DbType.Decimal, OTRate)
            DB.AddInParameter(DBC, "@FixedAllowance", DbType.Decimal, FixedAllowance)
            DB.AddInParameter(DBC, "@OtherAllowance", DbType.Decimal, OtherAllowance)
            DB.AddInParameter(DBC, "@Deductions", DbType.Decimal, Deductions)
            DB.AddInParameter(DBC, "@EPFNo", DbType.String, EPFNo)
            DB.AddInParameter(DBC, "@CreatedBy", DbType.Int64, CreatedBy)
            DB.AddInParameter(DBC, "@UpdatedBy", DbType.Int64, UpdatedBy)


            DB.ExecuteNonQuery(DBC)
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#Region "Employer Delete"
    Public Sub Delete()
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(EMPLOYERS_DELETE)
            DB.AddInParameter(DBC, "@EmployerID", DbType.Int64, Me.EmployerID)
            DB.ExecuteNonQuery(DBC)
        Catch ex As Exception
            Throw
        End Try

    End Sub

#End Region

#Region "Employer GetByID"
    Public Sub SelectRow()
        Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
        Dim DBC As DbCommand = DB.GetStoredProcCommand(EMPLOYERS_GETBYID)
        Try

            DB.AddInParameter(DBC, "@EmployerID", DbType.Int64, Me.EmployerID)
            Using DR As IDataReader = DB.ExecuteReader(DBC)


                With DR
                    Do While .Read
                        Me.EmployerID = Convert.ToInt64(IIf(Not IsDBNull(.Item("EmployerID")), Trim(.Item("EmployerID").ToString), 0))
                        Me.EmployerNo = Convert.ToInt64(IIf(Not IsDBNull(.Item("EmployerNo")), Trim(.Item("EmployerNo").ToString), 0))
                        Me.EmployerName = IIf(Not IsDBNull(.Item("EmployerName")), Trim(.Item("EmployerName").ToString), String.Empty)
                        Me.EmergencyContactPerson = IIf(Not IsDBNull(.Item("EmergencyContactPerson")), Trim(.Item("EmergencyContactPerson").ToString), String.Empty)

                        If Not IsDBNull(.Item("EmployerImage")) Then
                            EmployerImage = GetImageFromByteArray(.Item("EmployerImage"))
                        Else

                            EmployerImage = Nothing
                        End If

                        Me.AddressLine1 = IIf(Not IsDBNull(.Item("AddressLine1")), Trim(.Item("AddressLine1").ToString), String.Empty)
                        Me.AddressLine2 = IIf(Not IsDBNull(.Item("AddressLine2")), Trim(.Item("AddressLine2").ToString), String.Empty)
                        Me.AddressLine3 = IIf(Not IsDBNull(.Item("AddressLine3")), Trim(.Item("AddressLine3").ToString), String.Empty)

                        Me.DateOfBirth = Convert.ToDateTime((IIf(Not IsDBNull(.Item("DateOfBirth")), Trim(.Item("DateOfBirth").ToString), Date.MinValue)))
                        Me.DateJoined = Convert.ToDateTime((IIf(Not IsDBNull(.Item("DateJoined")), Trim(.Item("DateJoined").ToString), Date.MinValue)))

                        Me.Sex = IIf(Not IsDBNull(.Item("Sex")), Trim(.Item("Sex").ToString), String.Empty)
                        Me.NICNo = IIf(Not IsDBNull(.Item("NICNo")), Trim(.Item("NICNo").ToString), String.Empty)
                        Me.TelephoneNo = IIf(Not IsDBNull(.Item("TelephoneNo")), Trim(.Item("TelephoneNo").ToString), String.Empty)
                        Me.Designation = IIf(Not IsDBNull(.Item("Designation")), Trim(.Item("Designation").ToString), String.Empty)
                        Me.Department = IIf(Not IsDBNull(.Item("Department")), Trim(.Item("Department").ToString), String.Empty)
                        Me.BasicSalary = Convert.ToDecimal(IIf(Not IsDBNull(.Item("BasicSalary")), Trim(.Item("BasicSalary").ToString), 0))
                        Me.OTRate = Convert.ToDecimal(IIf(Not IsDBNull(.Item("OTRate")), Trim(.Item("OTRate").ToString), 0))
                        Me.FixedAllowance = Convert.ToDecimal(IIf(Not IsDBNull(.Item("FixedAllowance")), Trim(.Item("FixedAllowance").ToString), 0))
                        Me.OtherAllowance = Convert.ToDecimal(IIf(Not IsDBNull(.Item("OtherAllowance")), Trim(.Item("OtherAllowance").ToString), 0))
                        Me.EPFNo = IIf(Not IsDBNull(.Item("EPFNo")), Trim(.Item("EPFNo").ToString), String.Empty)
                        Me.Deductions = Convert.ToDecimal(IIf(Not IsDBNull(.Item("Deductions")), Trim(.Item("Deductions").ToString), 0))
                        Me.CreatedBy = Convert.ToInt64(IIf(Not IsDBNull(.Item("CreatedBy")), Trim(.Item("CreatedBy").ToString), 0))
                        Me.CreatedDate = Convert.ToDateTime((IIf(Not IsDBNull(.Item("CreatedDate")), Trim(.Item("CreatedDate").ToString), Date.MinValue)))
                        Me.UpdatedBy = Convert.ToInt64(IIf(Not IsDBNull(.Item("UpdatedBy")), Trim(.Item("UpdatedBy").ToString), 0))
                        Me.UpdatedDate = Convert.ToDateTime((IIf(Not IsDBNull(.Item("UpdatedDate")), Trim(.Item("UpdatedDate").ToString), Date.MinValue)))
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

#Region "Employer GetByEmployerNo"
    Public Sub GetByEmployerNo()
        Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
        Dim DBC As DbCommand = DB.GetStoredProcCommand(EMPLOYERS_GETBYEMPLOYERNO)
        Try

            DB.AddInParameter(DBC, "@EmployerNo", DbType.Int64, Me.EmployerNo)
            Using DR As IDataReader = DB.ExecuteReader(DBC)


                With DR
                    Do While .Read
                        Me.EmployerID = Convert.ToInt64(IIf(Not IsDBNull(.Item("EmployerID")), Trim(.Item("EmployerID").ToString), 0))
                        Me.EmployerNo = Convert.ToInt64(IIf(Not IsDBNull(.Item("EmployerNo")), Trim(.Item("EmployerNo").ToString), 0))
                        Me.EmployerName = IIf(Not IsDBNull(.Item("EmployerName")), Trim(.Item("EmployerName").ToString), String.Empty)
                        Me.EmergencyContactPerson = IIf(Not IsDBNull(.Item("EmergencyContactPerson")), Trim(.Item("EmergencyContactPerson").ToString), String.Empty)

                        If Not IsDBNull(.Item("EmployerImage")) Then
                            EmployerImage = GetImageFromByteArray(.Item("EmployerImage"))
                        Else

                            EmployerImage = Nothing
                        End If

                        Me.AddressLine1 = IIf(Not IsDBNull(.Item("AddressLine1")), Trim(.Item("AddressLine1").ToString), String.Empty)
                        Me.AddressLine2 = IIf(Not IsDBNull(.Item("AddressLine2")), Trim(.Item("AddressLine2").ToString), String.Empty)
                        Me.AddressLine3 = IIf(Not IsDBNull(.Item("AddressLine3")), Trim(.Item("AddressLine3").ToString), String.Empty)

                        Me.DateOfBirth = Convert.ToDateTime((IIf(Not IsDBNull(.Item("DateOfBirth")), Trim(.Item("DateOfBirth").ToString), Date.MinValue)))
                        Me.DateJoined = Convert.ToDateTime((IIf(Not IsDBNull(.Item("DateJoined")), Trim(.Item("DateJoined").ToString), Date.MinValue)))

                        Me.Sex = IIf(Not IsDBNull(.Item("Sex")), Trim(.Item("Sex").ToString), String.Empty)
                        Me.NICNo = IIf(Not IsDBNull(.Item("NICNo")), Trim(.Item("NICNo").ToString), String.Empty)
                        Me.TelephoneNo = IIf(Not IsDBNull(.Item("TelephoneNo")), Trim(.Item("TelephoneNo").ToString), String.Empty)
                        Me.Designation = IIf(Not IsDBNull(.Item("Designation")), Trim(.Item("Designation").ToString), String.Empty)
                        Me.Department = IIf(Not IsDBNull(.Item("Department")), Trim(.Item("Department").ToString), String.Empty)
                        Me.BasicSalary = Convert.ToDecimal(IIf(Not IsDBNull(.Item("BasicSalary")), Trim(.Item("BasicSalary").ToString), 0))
                        Me.OTRate = Convert.ToDecimal(IIf(Not IsDBNull(.Item("OTRate")), Trim(.Item("OTRate").ToString), 0))
                        Me.FixedAllowance = Convert.ToDecimal(IIf(Not IsDBNull(.Item("FixedAllowance")), Trim(.Item("FixedAllowance").ToString), 0))
                        Me.OtherAllowance = Convert.ToDecimal(IIf(Not IsDBNull(.Item("OtherAllowance")), Trim(.Item("OtherAllowance").ToString), 0))
                        Me.EPFNo = IIf(Not IsDBNull(.Item("EPFNo")), Trim(.Item("EPFNo").ToString), String.Empty)
                        Me.Deductions = Convert.ToDecimal(IIf(Not IsDBNull(.Item("Deductions")), Trim(.Item("Deductions").ToString), 0))
                        Me.CreatedBy = Convert.ToInt64(IIf(Not IsDBNull(.Item("CreatedBy")), Trim(.Item("CreatedBy").ToString), 0))
                        Me.CreatedDate = Convert.ToDateTime((IIf(Not IsDBNull(.Item("CreatedDate")), Trim(.Item("CreatedDate").ToString), Date.MinValue)))
                        Me.UpdatedBy = Convert.ToInt64(IIf(Not IsDBNull(.Item("UpdatedBy")), Trim(.Item("UpdatedBy").ToString), 0))
                        Me.UpdatedDate = Convert.ToDateTime((IIf(Not IsDBNull(.Item("UpdatedDate")), Trim(.Item("UpdatedDate").ToString), Date.MinValue)))
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

#Region "Getting Image from Byte Array"
    Public Function GetImageFromByteArray(ByVal bytes As Byte()) As Bitmap
        Try
            Return CType(Bitmap.FromStream(New IO.MemoryStream(bytes)), Bitmap)
        Catch ex As Exception
            Return Nothing
            Throw
        End Try


    End Function
#End Region

#Region "Employer SelectAll"
    Public Function SelectAll() As DataSet
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(EMPLOYERS_GETALL)
            Return DB.ExecuteDataSet(DBC)
            DBC.Dispose()
        Catch ex As Exception
            Return Nothing
            Throw
        End Try

    End Function
#End Region

#Region "Employer SelectRows By Designation"
    Public Function SelectRowsByDesignation() As DataSet
        Try
            Dim DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            Dim DBC As DbCommand = DB.GetStoredProcCommand(EMPLOYERS_GETBYBYDESIGNATION)
            DB.AddInParameter(DBC, "@Designation", DbType.String, Designation)
            Return DB.ExecuteDataSet(DBC)
            DBC.Dispose()
        Catch ex As Exception
            Return Nothing
            Throw
        End Try

    End Function
#End Region

End Class
