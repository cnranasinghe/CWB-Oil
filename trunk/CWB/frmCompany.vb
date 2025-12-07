Imports CWBCommon.CWBCompany
Imports CWBCommon.CWBConstants
Public Class frmCompany

#Region "Properties"
    Private _CWBCompany As CWBCommon.CWBCompany
#End Region

#Region "Constructors"
    
    Public ReadOnly Property CWBCompany() As CWBCommon.CWBCompany
        Get

            If _CWBCompany Is Nothing Then
                _CWBCompany = New CWBCommon.CWBCompany
            End If

            Return _CWBCompany
        End Get
    End Property
#End Region

#Region "Form Events"
    Private Sub frmCompany_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetStartUpBarButton(bbSave, bbNew, bbDelete, bbClose, bbDisplaySelected, bbRefresh, bbPrint)
        Me.HideToolButtonsOnLoad()
        Me.DisplayCompany()
    End Sub

    Private Sub frmCompany_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.Chr(27) Then
            Me.Close()
        End If
    End Sub
#End Region

#Region "Hide Tool Buttons On Load"
    Public Sub HideToolButtonsOnLoad()

        Me.bbDisplaySelected.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
        Me.bbRefresh.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
        Me.bbDelete.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
        Me.bbRefresh.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
        Me.bbNew.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
    End Sub
#End Region

#Region "Bar Button Events"
    Private Sub bbClose_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbClose.ItemClick
        Me.Close()
    End Sub

    Private Sub bbSave_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbSave.ItemClick

        If dxvpCompany.Validate Then
            Me.SaveCompany()
        End If


    End Sub
#End Region

#Region "Save Company"
    Private Sub SaveCompany()

        Try
            With CWBCompany
                .CompanyName = Me.teCompanyName.Text.Trim
                .Address1 = Me.teAddressLine1.Text.Trim
                .Address2 = Me.teAddressLine2.Text.Trim
                .Telephone = Me.teTelephone.Text.Trim
                .Fax = Me.teFax.Text.Trim
                .Email = Me.teEmail.Text.Trim
                .WebURL = Me.teWebURL.Text.Trim
                .RegNo = Me.teRegNo.Text.Trim
                .InsertCompany()
                Dim frm As New frmSavedOk
                frm.Text = CWB_SAVESUCCESS_CONFIRMATION_TITLE
                frm.lblTitle.Text = CWB_SAVESUCCESS_CONFIRMATION_TITLELABEL
                frm.lblDescription.Text = CWB_SAVESUCCESS_CONFIRMATION_DESCRIPTIONLABEL
                frm.ShowDialog()
            End With

        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
        



    End Sub
#End Region

#Region "Display Company"
    Private Sub DisplayCompany()
        Try
            With CWBCompany
                .GetCompany()
                Me.teCompanyName.Text = .CompanyName
                Me.teAddressLine1.Text = .Address1
                Me.teAddressLine2.Text = .Address2
                Me.teTelephone.Text = .Telephone
                Me.teFax.Text = .Fax
                Me.teEmail.Text = .Email
                Me.teWebURL.Text = .WebURL
                Me.teRegNo.Text = .RegNo
            End With
        Catch ex As Exception
            MessageError(ex.ToString)
        End Try

        
    End Sub
#End Region

End Class