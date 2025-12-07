Imports CWBCommon.CWBConstants
Public Class frmEmployers


#Region "Variables"
    Private _CWBEmployers As CWBCommon.CWBEmployers
#End Region

#Region "Constructors"


    Public ReadOnly Property CWBEmployers() As CWBCommon.CWBEmployers
        Get

            If _CWBEmployers Is Nothing Then
                _CWBEmployers = New CWBCommon.CWBEmployers
            End If

            Return _CWBEmployers
        End Get
    End Property
#End Region

#Region "Form Events"
    Private Sub frmEmployers_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.Chr(27) Then
            Me.Close()
        End If
    End Sub

    Private Sub frmEmployers_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetStartUpBarButton(bbSave, bbNew, bbDelete, bbClose, bbDisplaySelected, bbRefresh, bbPrint)
        Me.HideToolButtonsOnLoad()
    End Sub
#End Region

#Region "Tab Events"
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
        Me.bbPrint.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
        Me.bbSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Always
        Me.bbNew.Visibility = DevExpress.XtraBars.BarItemVisibility.Always
        Me.bbDelete.Visibility = DevExpress.XtraBars.BarItemVisibility.Always
    End Sub
#End Region

#Region "Bar Button Events"
    Private Sub bbSave_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbSave.ItemClick
        If dxvpEmployers.Validate Then
            If lblID.Text = String.Empty Then

                CWBEmployers.EmployerNo = Me.seEmployerNo.Text.Trim

                If CWBEmployers.IsEmployerExits() = True Then
                    Dim frm As New frmAlreadyExists

                    frm.Text = CWB_EMPLOYER_CONFIRMATION_TITLE
                    frm.lblTitle.Text = CWB_EMPLOYER_CONFIRMATION_TITLELABEL
                    frm.lblDescription.Text = CWB_EMPLOYER_CONFIRMATION_DESCRIPTIONLABEL
                    frm.ShowDialog()
                    Me.seEmployerNo.Focus()
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

                    CWBEmployers.EmployerNo = Me.seEmployerNo.Text
                    CWBEmployers.GetByEmployerNo()

                    If CWBEmployers.EmployerID = Me.lblID.Text Then
                        Me.SaveRecords()
                    Else

                        Dim frmx As New frmAlreadyExists

                        frmx.Text = CWB_EMPLOYER_CONFIRMATION_TITLE
                        frmx.lblTitle.Text = CWB_EMPLOYER_CONFIRMATION_TITLELABEL
                        frmx.lblDescription.Text = CWB_EMPLOYER_CONFIRMATION_DESCRIPTIONLABEL
                        frmx.ShowDialog()
                        Me.seEmployerNo.Focus()

                    End If

                End If
            End If

        End If
    End Sub

    Private Sub bbClose_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbClose.ItemClick
        Me.Close()
    End Sub

    Private Sub bbNew_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbNew.ItemClick
        Me.ClearFormData()
    End Sub

    Private Sub bbDelete_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbDelete.ItemClick
        If Not lblID.Text = String.Empty Then
            Dim frm As New frmDeleteYesNo
            frm.lblTitle.Text = CWB_DELETE_CONFIRMATION_TITLELABEL
            frm.lblDescription.ForeColor = Color.Red
            frm.peImage.Image = CWB.My.Resources.Resources.ImgDelete
            frm.lblDescription.Text = CWB_DELETE_CONFIRMATION_DESCRIPTIONLABEL
            frm.Text = CWB_DELETE_CONFIRMATION_TITLE

            If frm.ShowDialog = Windows.Forms.DialogResult.Yes Then
                Me.Delete(lblID.Text)

            End If
        End If

    End Sub

    Private Sub bbRefresh_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbRefresh.ItemClick
        With gvEmployers
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

        PrintPreview(gcEmployers, "Employees Report")
    End Sub
#End Region

#Region "Grid Events"

    Private Sub gvEmployers_DoubleClick_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gvEmployers.DoubleClick
        Me.DisplayRecord()
    End Sub
#End Region

#Region "Save Records"
    Private Sub SaveRecords()





        Try
            With CWBEmployers
                .EmployerID = Convert.ToInt64(IIf(lblID.Text = String.Empty, 0, lblID.Text))
                .EmployerNo = seEmployerNo.Text
                .EmployerName = teEmployerName.Text.Trim
                .EmployerImage = peEmployer.Image
                .AddressLine1 = teAdressLine1.Text.Trim
                .AddressLine2 = teAddressLine2.Text.Trim
                .AddressLine3 = teCity.Text.Trim
                If IsDate(deDOB.EditValue) Then
                    .DateOfBirth = deDOB.EditValue
                Else
                    .DateOfBirth = Date.MinValue
                End If

                If IsDate(deDateJoined.EditValue) Then
                    .DateJoined = deDateJoined.EditValue
                Else
                    .DateJoined = Date.MinValue
                End If

                .Sex = ceSex.Text
                .NICNo = teNICNo.Text.Trim
                .TelephoneNo = teTelephone.Text.Trim
                .Designation = ceDesignation.Text
                .Department = ceDepartment.Text
                .EmergencyContactPerson = teEmployerContactPerson.Text.Trim
                .BasicSalary = seBasicSalary.Text
                .OTRate = seOTRate.Text
                .FixedAllowance = seFixedAllowance.Text
                .OtherAllowance = seOtherAllowance.Text
                .Deductions = seDeduction.Text
                .EPFNo = teEPFNo.Text.Trim

                .CreatedBy = UserID
                .UpdatedBy = UserID
                .Insert()
                Dim frm As New frmSavedOk
                frm.Text = CWB_SAVESUCCESS_CONFIRMATION_TITLE
                frm.lblTitle.Text = CWB_SAVESUCCESS_CONFIRMATION_TITLELABEL
                frm.lblDescription.Text = CWB_SAVESUCCESS_CONFIRMATION_DESCRIPTIONLABEL
                frm.ShowDialog()
            End With
            Me.ClearFormData()
        Catch ex As Exception
            MessageError(ex.ToString)

        End Try

    End Sub
#End Region

#Region "Delete Data"
    Private Sub Delete(ByVal EmployerID As Int64)

        Try
            With CWBEmployers
                .EmployerID = EmployerID
                .Delete()
            End With
            Me.ClearFormData()

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
            If gvEmployers.RowCount > 0 And Not gvEmployers.IsFilterRow(gvEmployers.FocusedRowHandle) Then
                xTab1.SelectedTabPageIndex = 0
                With CWBEmployers

                    .EmployerID = Me.gvEmployers.GetFocusedRowCellValue(GridColumn1)
                    .SelectRow()

                    lblID.Text = .EmployerID

                    seEmployerNo.Text = .EmployerNo
                    teEmployerName.Text = .EmployerName
                    peEmployer.EditValue = .EmployerImage
                    teAdressLine1.Text = .AddressLine1
                    teAddressLine2.Text = .AddressLine2
                    teCity.Text = .AddressLine3

                    If .DateOfBirth = Date.MinValue Then
                        deDOB.EditValue = DBNull.Value
                    Else
                        deDOB.EditValue = .DateOfBirth
                    End If


                    If .DateJoined = Date.MinValue Then
                        deDateJoined.EditValue = DBNull.Value

                    Else
                        deDateJoined.EditValue = .DateJoined

                    End If


                    ceSex.Text = .Sex
                    teNICNo.Text = .NICNo
                    teTelephone.Text = .TelephoneNo
                    ceDesignation.Text = .Designation
                    ceDepartment.Text = .Department
                    teEmployerContactPerson.Text = .EmergencyContactPerson
                    seBasicSalary.Text = .BasicSalary
                    seOTRate.Text = .OTRate
                    seFixedAllowance.Text = .FixedAllowance
                    seOtherAllowance.Text = .OtherAllowance
                    seDeduction.Text = .Deductions
                    teEPFNo.Text = .EPFNo


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
            gcEmployers.DataSource = CWBEmployers.SelectAll.Tables(0)
        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Clears Form Data"

    Private Sub ClearFormData()
        Try
            lblID.Text = String.Empty
            seEmployerNo.Text = String.Empty
            teEmployerName.Text = String.Empty
            peEmployer.EditValue = DBNull.Value
            teAdressLine1.Text = String.Empty
            teAddressLine2.Text = String.Empty
            teCity.Text = String.Empty
            deDOB.EditValue = DBNull.Value
            deDateJoined.EditValue = DBNull.Value
            ceSex.Text = String.Empty
            teNICNo.Text = String.Empty
            teTelephone.Text = String.Empty
            ceDesignation.Text = String.Empty
            ceDepartment.Text = String.Empty
            teEmployerContactPerson.Text = String.Empty
            seBasicSalary.Text = 0
            seOTRate.Text = 0
            seFixedAllowance.Text = 0
            seOtherAllowance.Text = 0.0
            seDeduction.Text = 0
            teEPFNo.Text = String.Empty
            dxvpEmployers.RemoveControlError(seEmployerNo)
            dxvpEmployers.RemoveControlError(teEmployerName)
            dxvpEmployers.RemoveControlError(ceSex)


            seEmployerNo.Focus()
        Catch ex As Exception
            MessageError(ex.ToString)
        End Try

    End Sub

#End Region


End Class