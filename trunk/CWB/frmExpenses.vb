Imports CWBCommon.CWBExpenses
Imports CWBCommon.CWBCollections
Imports CWBCommon.CWBConstants
Imports Microsoft.Practices.EnterpriseLibrary.Common
Imports Microsoft.Practices.EnterpriseLibrary.Data
Imports System.Data
Imports System.Data.Common
Imports CWBCommon.CWBEnums


Public Class frmExpenses

#Region "Variables"
    Private _CWBExpense As CWBCommon.CWBExpenses
    Private _CWBCollection As CWBCommon.CWBCollections
    Dim _CWBEnums As CWBCommon.CWBEnums
#End Region

#Region "Constructors"
    
    Public ReadOnly Property CWBExpense() As CWBCommon.CWBExpenses
        Get

            If _CWBExpense Is Nothing Then
                _CWBExpense = New CWBCommon.CWBExpenses()
            End If

            Return _CWBExpense
        End Get
    End Property
    Public ReadOnly Property CWBEnums() As CWBCommon.CWBEnums
        Get

            If _CWBEnums Is Nothing Then
                _CWBEnums = New CWBCommon.CWBEnums
            End If

            Return _CWBEnums
        End Get
    End Property


    Public ReadOnly Property CWBCollection() As CWBCommon.CWBCollections
        Get

            If _CWBCollection Is Nothing Then
                _CWBCollection = New CWBCommon.CWBCollections
            End If

            Return _CWBCollection
        End Get
    End Property
#End Region

#Region "Form Events"
    Private Sub frmExpenses_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.Chr(27) Then
            Me.Close()
        End If
    End Sub

    Private Sub frmExpenses_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetCWBExpenseBarButton(bbSave, bbNew, bbDelete, bbClose, bbDisplaySelected, bbRefresh, bbPrint)
        Me.HideToolButtonsOnLoad()
        Me.PopulateExpenseTypesLookup()
        Me.deDate.EditValue = Date.Today
        Me.deFromDate.EditValue = Date.Today
        Me.deToDate.EditValue = Date.Today
    End Sub
#End Region

#Region "Bar Button Events"

    Private Sub bbClose_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbClose.ItemClick
        Me.Close()
    End Sub

    Private Sub bbNew_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbNew.ItemClick
        Me.ClearFormData()
    End Sub

    Private Sub bbSave_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbSave.ItemClick

        If dxvpExpenses.Validate Then
            If lblID.Text = String.Empty Then
                Me.SaveRecords()
            Else
                Dim frm As New frmUpdateYesNo
                frm.peImage.Image = CWB.My.Resources.Resources.ImgUpdate
                frm.Text = CWB_UPDATE_CONFIRMATION_TITLE
                frm.lblTitle.Text = CWB_UPDATE_CONFIRMATION_TITLELABEL
                frm.lblDescription.Text = CWB_UPDATE_CONFIRMATION_DESCRIPTIONLABEL


                If frm.ShowDialog = Windows.Forms.DialogResult.Yes Then
                    Me.UpdateRecords()

                End If
            End If
            Me.ClearFormData()
        End If
    End Sub

    Private Sub bbRefresh_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbRefresh.ItemClick
        With gvExpenses
            .ClearColumnsFilter()
            .ClearGrouping()
            .ClearSelection()
            .ClearSorting()
        End With
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

                Me.DeleteExpenses(lblID.Text)
                Me.ClearFormData()
            End If
        End If

        lupExpenseType.Focus()

    End Sub

    Private Sub bbDisplaySelected_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbDisplaySelected.ItemClick
        Me.DisplayExpense()
    End Sub

    Private Sub bbPrint_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbPrint.ItemClick
        PrintPreview(gcExpenses, "Expense Report")
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
        Me.bbSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
        Me.bbNew.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
        Me.bbDelete.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
        Me.bbPrint.Visibility = DevExpress.XtraBars.BarItemVisibility.Always

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
                

        End Select

    End Sub
#End Region

#Region "Clear Form Data"
    Private Sub ClearFormData()
        Me.lupExpenseType.EditValue = DBNull.Value
        Me.seAmount.EditValue = 0
        Me.meNote.Text = String.Empty
        Me.lblID.Text = String.Empty
        dxvpExpenses.RemoveControlError(lupExpenseType)
        Me.lupExpenseType.Focus()

    End Sub
#End Region

#Region "Populate ExpenseTypes Lookup"
    Public Sub PopulateExpenseTypesLookup()

        Try
            With lupExpenseType
                .Properties.DataSource = CWBExpense.ExpenseTypesGetAll().Tables(0)
                .Properties.DisplayMember = "Description"
                .Properties.ValueMember = "ExpenseTypeID"

            End With


        Catch ex As Exception

        End Try
    End Sub
#End Region

#Region "Lookup Events"
    Private Sub lupExpenseType_ButtonClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraEditors.Controls.ButtonPressedEventArgs) Handles lupExpenseType.ButtonClick
        Select Case e.Button.Index
            Case 1
                If Not Me.lupExpenseType.Text = String.Empty Then
                    Me.SaveExpenseType()
                    Me.PopulateExpenseTypesLookup()
                End If

            Case 2


                If Not Me.lupExpenseType.Text = String.Empty Then
                    Dim frm As New frmDeleteYesNo
                    frm.lblTitle.Text = CWB_DELETE_CONFIRMATION_TITLELABEL
                    frm.lblDescription.ForeColor = Color.Red
                    frm.peImage.Image = CWB.My.Resources.Resources.ImgDelete
                    frm.lblDescription.Text = CWB_DELETE_CONFIRMATION_DESCRIPTIONLABEL
                    frm.Text = CWB_DELETE_CONFIRMATION_TITLE

                    If frm.ShowDialog = Windows.Forms.DialogResult.Yes Then
                        Me.DeleteExpenseType()
                        Me.PopulateExpenseTypesLookup()

                    End If
                End If

        End Select

    End Sub
#End Region

#Region "Save Expenses"
    Private Sub SaveRecords()

        Dim _Connection As DbConnection = Nothing
        Dim _Transaction As DbTransaction = Nothing

        Try

            Dim _DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            _Connection = _DB.CreateConnection
            _Connection.Open()
            _Transaction = _Connection.BeginTransaction()

            With CWBExpense
                .ExpenseID = Convert.ToInt64(IIf(lblID.Text = String.Empty, 0, lblID.Text))
                .ExpenseTypeID = Me.lupExpenseType.EditValue
                Select Case cbePaymentType.SelectedIndex
                    Case 0
                        .PaymentTypeID = CWBCommon.CWBEnums.EnumPaymentTypes.CASH
                    Case 1
                        .PaymentTypeID = CWBCommon.CWBEnums.EnumPaymentTypes.CHECK
                    Case 2
                        .PaymentTypeID = CWBCommon.CWBEnums.EnumPaymentTypes.CCARD
                End Select

                .ExpenseDate = Me.deDate.EditValue
                .Amount = Me.seAmount.EditValue
                .Note = Me.meNote.EditValue
                .CreatedBy = UserID
                .UpdatedBy = UserID
                .InsertExpnese(_DB, _Transaction)

                With CWBCollection
                    .SystemID = CWBExpense.CurrentExpenseID
                    .TransactionTypeID = CWBCommon.CWBEnums.EnumTransactionTypes.EXPENSES

                    Select Case Me.cbePaymentType.SelectedIndex
                        Case 0
                            .PaymentTypeID = CWBCommon.CWBEnums.EnumPaymentTypes.CASH
                        Case 1
                            .PaymentTypeID = CWBCommon.CWBEnums.EnumPaymentTypes.CHECK
                        Case 2
                            .PaymentTypeID = CWBCommon.CWBEnums.EnumPaymentTypes.CCARD
                    End Select

                    .Date = Me.deDate.EditValue
                    .Reference = String.Empty
                    .Amount = Me.seAmount.EditValue
                    .InsertCollection(_DB, _Transaction)

                End With
                _Transaction.Commit()
                Dim frm As New frmSavedOk
                frm.Text = CWB_SAVESUCCESS_CONFIRMATION_TITLE
                frm.lblTitle.Text = CWB_SAVESUCCESS_CONFIRMATION_TITLELABEL
                frm.lblDescription.Text = CWB_SAVESUCCESS_CONFIRMATION_DESCRIPTIONLABEL
                frm.ShowDialog()
            End With
        Catch ex As Exception
            _Transaction.Rollback()
        End Try

    End Sub
#End Region

#Region "Update Expenses"
    Private Sub UpdateRecords()

        Dim _Connection As DbConnection = Nothing
        Dim _Transaction As DbTransaction = Nothing

        Try

            Dim _DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            _Connection = _DB.CreateConnection
            _Connection.Open()
            _Transaction = _Connection.BeginTransaction()

            With CWBExpense

                .ExpenseID = Convert.ToInt64(IIf(lblID.Text = String.Empty, 0, lblID.Text))
                .ExpenseTypeID = Me.lupExpenseType.EditValue

                Select Case cbePaymentType.SelectedIndex
                    Case 0
                        .PaymentTypeID = CWBCommon.CWBEnums.EnumPaymentTypes.CASH
                    Case 1
                        .PaymentTypeID = CWBCommon.CWBEnums.EnumPaymentTypes.CHECK
                    Case 2
                        .PaymentTypeID = CWBCommon.CWBEnums.EnumPaymentTypes.CCARD
                End Select

                .ExpenseDate = Me.deDate.EditValue
                .Amount = Me.seAmount.EditValue
                .Note = Me.meNote.EditValue
                .CreatedBy = UserID
                .UpdatedBy = UserID
                .InsertExpnese(_DB, _Transaction)



                With CWBCollection

                    .SystemID = Convert.ToInt64(IIf(lblID.Text = String.Empty, 0, lblID.Text))
                    .TransactionTypeID = CWBCommon.CWBEnums.EnumTransactionTypes.EXPENSES
                    .CollectionDelete(_DB, _Transaction)

                    .SystemID = Convert.ToInt64(IIf(lblID.Text = String.Empty, 0, lblID.Text))
                    .TransactionTypeID = CWBCommon.CWBEnums.EnumTransactionTypes.EXPENSES

                    Select Case Me.cbePaymentType.SelectedIndex
                        Case 0
                            .PaymentTypeID = CWBCommon.CWBEnums.EnumPaymentTypes.CASH
                        Case 1
                            .PaymentTypeID = CWBCommon.CWBEnums.EnumPaymentTypes.CHECK
                        Case 2
                            .PaymentTypeID = CWBCommon.CWBEnums.EnumPaymentTypes.CCARD
                    End Select

                    .Date = Date.Now
                    .Reference = String.Empty
                    .Amount = Me.seAmount.EditValue
                    .InsertCollection(_DB, _Transaction)

                End With
                _Transaction.Commit()
                Dim frm As New frmSavedOk
                frm.Text = CWB_SAVESUCCESS_CONFIRMATION_TITLE
                frm.lblTitle.Text = CWB_SAVESUCCESS_CONFIRMATION_TITLELABEL
                frm.lblDescription.Text = CWB_SAVESUCCESS_CONFIRMATION_DESCRIPTIONLABEL
                frm.ShowDialog()
            End With
        Catch ex As Exception
            _Transaction.Rollback()
        End Try

    End Sub
#End Region

#Region "Save Expense Type"
    Private Sub SaveExpenseType()
        Try
            With CWBExpense
                .Description = Me.lupExpenseType.Text
                .InsertExpneseTypes()

            End With
        Catch ex As Exception

        End Try
    End Sub
#End Region

#Region "Delete Expense Type"
    Private Sub DeleteExpenseType()
        Try
            With CWBExpense
                .Description = Me.lupExpenseType.Text
                .ExpenseTypesDelete()
            End With
        Catch ex As Exception

        End Try
    End Sub
#End Region

#Region "Populate Grid"
    Public Sub PopulateHistoryData()
        Try
            CWBExpense.FromDate = Me.deFromDate.EditValue
            CWBExpense.ToDate = Me.deToDate.EditValue

            gcExpenses.DataSource = CWBExpense.ExpenseGetByDates().Tables(0)
        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Delete Expenses"
    Private Sub DeleteExpenses(ByVal ExpenseID As Int64)
        Try
            With CWBExpense
                .ExpenseID = ExpenseID
                .ExpenseDelete()
            End With

        Catch ex As Exception
            Dim frm As New frmOk
            frm.Text = CWB_DELETE_ERROR_CONFIRMATION_TITLE
            frm.lblTitle.Text = CWB_DELETE_ERROR_CONFIRMATION_TITLELABEL
            frm.lblDescription.Text = CWB_DELETE_ERROR_CONFIRMATION_DESCRIPTIONLABEL
            frm.ShowDialog()
        End Try
        

    End Sub
#End Region

#Region "Display Expense"
    Public Sub DisplayExpense()

        If gvExpenses.RowCount > 0 Then


            xTab1.SelectedTabPageIndex = 0
            With CWBExpense
                .ExpenseID = gvExpenses.GetFocusedRowCellValue(GridColumn1)
                .GetExpenseByID()
                lblID.Text = .ExpenseID
                lupExpenseType.EditValue = .ExpenseTypeID

                Select Case .PaymentTypeID
                    Case 1
                        cbePaymentType.SelectedIndex = 0
                    Case 2
                        cbePaymentType.SelectedIndex = 1
                    Case 3
                        cbePaymentType.SelectedIndex = 2


                End Select

                deDate.EditValue = .ExpenseDate
                seAmount.EditValue = .Amount
                meNote.EditValue = .Note


            End With
        End If

    End Sub
#End Region

#Region "Grid Events"
    Private Sub gvExpenses_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gvExpenses.DoubleClick
        Me.DisplayExpense()
    End Sub
#End Region

#Region "Button Events"
    Private Sub sbProcess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sbProcess.Click
        Try

            If dxvpHistoryData.Validate Then
                Me.PopulateHistoryData()
            End If

        Catch ex As Exception
            MessageError(ex.ToString)

        End Try

    End Sub
#End Region

End Class