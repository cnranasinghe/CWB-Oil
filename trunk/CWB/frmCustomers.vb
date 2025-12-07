Imports CWBCommon.CWBCustomers
Imports CWBCommon.CWBConstants
Public Class frmCustomers

#Region "Variables"
    Private _CWBCustomers As CWBCommon.CWBCustomers
#End Region

#Region "Constructors"
    Public ReadOnly Property CWBCustomers() As CWBCommon.CWBCustomers
        Get
            If _CWBCustomers Is Nothing Then
                _CWBCustomers = New CWBCommon.CWBCustomers
            End If

            Return _CWBCustomers
        End Get
    End Property
#End Region

#Region "Form Events"
    Private Sub frmCustomers_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.Chr(27) Then
            Me.Close()
        End If
    End Sub

    Private Sub frmCustomers_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetStartUpBarButton(bbSave, bbNew, bbDelete, bbClose, bbDisplaySelected, bbRefresh, bbPrint)

        Me.HideToolButtonsOnLoad()
    End Sub
#End Region

#Region "Hide Tool Buttons On Load"
    Public Sub HideToolButtonsOnLoad()

        Me.bbDisplaySelected.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
        Me.bbRefresh.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
        Me.bbPrint.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
    End Sub
#End Region

#Region "Show Tool Buttons On History Tab change"
    Public Sub ShowToolButtonsOnHistoryTabChange()

        Me.bbDisplaySelected.Visibility = DevExpress.XtraBars.BarItemVisibility.Always
        Me.bbRefresh.Visibility = DevExpress.XtraBars.BarItemVisibility.Always
        Me.bbPrint.Visibility = DevExpress.XtraBars.BarItemVisibility.Always
        Me.bbSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
        Me.bbNew.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
        Me.bbDelete.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
    End Sub
#End Region

#Region "Show Tool Buttons On New Record Tab change"
    Public Sub ShowToolButtonsOnNewRecordTabChange()

        Me.bbDisplaySelected.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
        Me.bbRefresh.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
        Me.bbSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Always
        Me.bbNew.Visibility = DevExpress.XtraBars.BarItemVisibility.Always
        Me.bbDelete.Visibility = DevExpress.XtraBars.BarItemVisibility.Always
        Me.bbPrint.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
    End Sub
#End Region

#Region "Tab Control Events"
    Private Sub xTab1_SelectedPageChanged(ByVal sender As System.Object, ByVal e As DevExpress.XtraTab.TabPageChangedEventArgs) Handles xTab1.SelectedPageChanged
        Select Case e.Page.TabControl.SelectedTabPageIndex
            Case 0
                Me.ShowToolButtonsOnNewRecordTabChange()
            Case 1
                Me.ShowToolButtonsOnHistoryTabChange()
                Me.PopulateGrid()

        End Select
    End Sub
#End Region

#Region "Grid Events"
    Private Sub gvCustomers_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gvCustomers.DoubleClick
        Me.DisplayRecord()
    End Sub
#End Region

#Region "Bar Button Events"
    Private Sub bbSave_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbSave.ItemClick
        If dxvpCustomers.Validate Then
            If lblCustomerID.Text = String.Empty Then




                CWBCustomers.CustomerNo = Me.seCustomerNo.Text

                If CWBCustomers.IsCustomerExits() = True Then
                    Dim frm As New frmAlreadyExists

                    frm.Text = CWB_CUSTOMER_CONFIRMATION_DESCRIPTIONLABEL
                    frm.lblTitle.Text = CWB_CUSTOMER_CONFIRMATION_TITLELABEL
                    frm.lblDescription.Text = CWB_CUSTOMER_CONFIRMATION_DESCRIPTIONLABEL
                    frm.ShowDialog()
                    Me.seCustomerNo.Focus()
                Else

                    Me.SaveRecords()
                End If


            Else
                Dim frm As New frmUpdateYesNo
                frm.peImage.Image = CWB.My.Resources.Resources.ImgUpdate
                frm.Text = CWB_UPDATE_CONFIRMATION_TITLE
                frm.lblTitle.Text = CWB_UPDATE_CONFIRMATION_TITLELABEL
                frm.lblDescription.Text = CWB_UPDATE_CONFIRMATION_DESCRIPTIONLABEL

              

                If frm.ShowDialog = Windows.Forms.DialogResult.Yes Then

                    CWBCustomers.CustomerNo = Me.seCustomerNo.Text
                    CWBCustomers.GetCustomerByCustomerNo()

                    If CWBCustomers.CustomerID = Me.lblCustomerID.Text Then
                        Me.SaveRecords()
                    Else

                        Dim frmx As New frmAlreadyExists

                        frmx.Text = CWB_SUPPLIER_CONFIRMATION_TITLE
                        frmx.lblTitle.Text = CWB_SUPPLIER_CONFIRMATION_TITLELABEL
                        frmx.lblDescription.Text = CWB_SUPPLIER_CONFIRMATION_DESCRIPTIONLABEL
                        frmx.ShowDialog()
                        Me.seCustomerNo.Focus()

                    End If

                End If

            End If

        End If
    End Sub

    Private Sub bbClose_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbClose.ItemClick
        Me.Close()
    End Sub

    Private Sub bbNew_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbNew.ItemClick
        Me.ClearFromData()
    End Sub

    Private Sub bbDelete_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbDelete.ItemClick
        If Not lblCustomerID.Text = String.Empty Then
            Dim frm As New frmDeleteYesNo
            frm.lblTitle.Text = CWB_DELETE_CONFIRMATION_TITLELABEL
            frm.lblDescription.ForeColor = Color.Red
            frm.peImage.Image = CWB.My.Resources.Resources.ImgDelete
            frm.lblDescription.Text = CWB_DELETE_CONFIRMATION_DESCRIPTIONLABEL
            frm.Text = CWB_DELETE_CONFIRMATION_TITLE

            If frm.ShowDialog = Windows.Forms.DialogResult.Yes Then
                Me.DeleteCustomer(lblCustomerID.Text)

            End If
        End If

    End Sub

    Private Sub bbRefresh_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbRefresh.ItemClick
        With gvCustomers
            .ClearColumnsFilter()
            .ClearGrouping()
            .ClearSelection()
            .ClearSorting()
        End With

    End Sub

    Private Sub bbDisplaySelected_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbDisplaySelected.ItemClick
        Me.DisplayRecord()
    End Sub

    Private Sub bbPrint_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbPrint.ItemClick
        PrintPreview(gcCustomers, "Customer Report")

    End Sub
#End Region

#Region "Save Customer Data"
    Private Sub SaveRecords()
        Try
            With CWBCustomers
                .CustomerID = Convert.ToInt64(IIf(lblCustomerID.Text = String.Empty, 0, lblCustomerID.Text))
                .CustomerNo = seCustomerNo.Text
                .Salutation = cbeSalutation.Text
                .CustomerName = teCustomerName.Text.Trim
                .AddressLine1 = teAddressLine1.Text.Trim
                .AddressLine2 = teAddressLine2.Text.Trim
                .AddressLine3 = teAddressLine3.Text.Trim
                .Telephone = teTelephone.Text.Trim
                .Fax = teFax.Text.Trim
                .Email = teEmail.Text.Trim
                .WebURL = teWebURL.Text.Trim
                .CreatedBy = UserID
                .UpdatedBy = UserID
                .InsertCustomers()
                Dim frm As New frmSavedOk
                frm.Text = CWB_SAVESUCCESS_CONFIRMATION_TITLE
                frm.lblTitle.Text = CWB_SAVESUCCESS_CONFIRMATION_TITLELABEL
                frm.lblDescription.Text = CWB_SAVESUCCESS_CONFIRMATION_DESCRIPTIONLABEL
                frm.ShowDialog()
            End With
            Me.ClearFromData()
        Catch ex As Exception
            MessageError(ex.ToString)
        End Try

    End Sub
#End Region

#Region "Delete Customer"
    Private Sub DeleteCustomer(ByVal CustomerID As Int64)

        Try
            With CWBCustomers
                .CustomerID = CustomerID
                .DeleteCustomer()
            End With
            Me.ClearFromData()
        Catch ex As SqlClient.SqlException
            Dim frm As New frmOk
            frm.Text = CWB_DELETE_ERROR_CONFIRMATION_TITLE
            frm.lblTitle.Text = CWB_DELETE_ERROR_CONFIRMATION_TITLELABEL
            frm.lblDescription.Text = CWB_DELETE_ERROR_CONFIRMATION_DESCRIPTIONLABEL
            frm.ShowDialog()
        End Try

    End Sub
#End Region

#Region "Dispay Record"
    Public Sub DisplayRecord()

        Try
            If gvCustomers.RowCount > 0 And Not gvCustomers.IsFilterRow(gvCustomers.FocusedRowHandle) Then
                xTab1.SelectedTabPageIndex = 0
                With CWBCustomers

                    .CustomerID = Me.gvCustomers.GetFocusedRowCellValue(GridColumn1)
                    .GetCustomerByID()
                    lblCustomerID.Text = .CustomerID
                    seCustomerNo.Text = .CustomerNo
                    cbeSalutation.Text = .Salutation
                    teCustomerName.Text = .CustomerName
                    teAddressLine1.Text = .AddressLine1
                    teAddressLine2.Text = .AddressLine2
                    teAddressLine3.Text = .AddressLine3
                    teTelephone.Text = .Telephone
                    teFax.Text = .Fax
                    teEmail.Text = .Email
                    teWebURL.EditValue = .WebURL

                End With


            End If
        Catch ex As Exception
            MessageError(ex.ToString)
        End Try

    End Sub
#End Region

#Region "Populate Grid"
    Public Sub PopulateGrid()
        Try
            gcCustomers.DataSource = CWBCustomers.GetAllCustomers.Tables(0)
        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Clears Form Data"

    Private Sub ClearFromData()
        Try
            lblCustomerID.Text = String.Empty
            seCustomerNo.Text = String.Empty
            cbeSalutation.Text = String.Empty
            teCustomerName.Text = String.Empty
            teAddressLine1.Text = String.Empty
            teAddressLine2.Text = String.Empty
            teAddressLine3.Text = String.Empty
            teTelephone.Text = String.Empty
            teFax.Text = String.Empty
            teEmail.Text = String.Empty
            teWebURL.Text = String.Empty
            dxvpCustomers.RemoveControlError(seCustomerNo)
            dxvpCustomers.RemoveControlError(teCustomerName)
            seCustomerNo.Focus()

        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
        
    End Sub

#End Region

   
   
  
End Class