
Public Class frmLogin




#Region "Variables"
    Private _CWBUserLogins As CWBCommon.CWBUserLogins
    Private _CWBLogBook As CWBCommon.CWBLogBook

#End Region

#Region "Constructors"
    Public ReadOnly Property CWBUserLogins() As CWBCommon.CWBUserLogins
        Get

            If _CWBUserLogins Is Nothing Then
                _CWBUserLogins = New CWBCommon.CWBUserLogins()
            End If

            Return _CWBUserLogins
        End Get
    End Property

   
    Public ReadOnly Property CWBLogBook() As CWBCommon.CWBLogBook
        Get
            ' Create on demand...
            If _CWBLogBook Is Nothing Then
                _CWBLogBook = New CWBCommon.CWBLogBook()
            End If

            Return _CWBLogBook
        End Get
    End Property
#End Region

#Region "Form Events"
    Private Sub frmLogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DevExpress.Skins.SkinManager.EnableFormSkins()
        DevExpress.LookAndFeel.LookAndFeelHelper.ForceDefaultLookAndFeelChanged()
        Me.PopulateLookup()
    End Sub
#End Region

#Region "Populate Users"
    Public Sub PopulateLookup()
        Try
            leUsers.Properties.DataSource = CWBUserLogins.SelectAll.Tables(0)
            leUsers.Properties.DisplayMember = "UserName"
            leUsers.Properties.ValueMember = "UserLoginID"

        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Button Events"
    Private Sub SimpleButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sbExit.Click
        Application.Exit()
    End Sub

    Private Sub sbLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sbLogin.Click

        Try
            If dxvpLogin.Validate Then


                With CWBLogBook
                    .UserName = Me.leUsers.Text
                    .Password = Me.tePassword.Text
                    .IPAddress = "Assing By DHCP"
                    .ComputerName = Environment.MachineName
                    .InsertLogbook()

                End With



                CWBUserLogins.UserName = Me.leUsers.Text
                CWBUserLogins.Password = Me.tePassword.Text.Trim

                CWBUserLogins.SelectRowByUserNameAndPassword()

                If IsDBNull(CWBUserLogins.UserLoginID) Or CWBUserLogins.UserLoginID = 0 Then
                    lblError.Text = "Wrong Password"
                Else
                    lblError.Text = "Athunticated"

                    UserID = CWBUserLogins.UserLoginID
                    UserName = CWBUserLogins.UserName
                    UserType = CWBUserLogins.UserType


                    Select Case UserType

                        Case "USER"
                            frmMain.RibbonPage4.Visible = False
                            frmMain.RibbonPageGroup9.Visible = False
                            frmMain.BarButtonItem17.Visibility = DevExpress.XtraBars.BarItemVisibility.Never


                    End Select

                    Me.Hide()
                    frmMain.Show()


                End If

            Else
                tePassword.Focus()
            End If

        Catch ex As Exception
            MessageError(ex.ToString)
        End Try



    End Sub
#End Region


End Class