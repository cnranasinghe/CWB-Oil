Imports CWBCommon.CWBSales
Imports CWBCommon.CWBCollections
Imports Microsoft.Practices.EnterpriseLibrary.Common
Imports Microsoft.Practices.EnterpriseLibrary.Data
Imports System.Data
Imports System.Data.Common
Imports CWBCommon.CWBConstants
Imports CWBCommon.CWBEnums


Public Class frmReceivePayments

#Region "Variables"
    Dim _DASales As CWBCommon.CWBSales
    Dim _CWBCollections As CWBCommon.CWBCollections
    Dim _CWBEnums As CWBCommon.CWBEnums
#End Region

#Region "Constructor"


    Public ReadOnly Property CWBSales() As CWBCommon.CWBSales
        Get

            If _DASales Is Nothing Then
                _DASales = New CWBCommon.CWBSales
            End If

            Return _DASales
        End Get
    End Property


    Public ReadOnly Property CWBCollections() As CWBCommon.CWBCollections
        Get

            If _CWBCollections Is Nothing Then
                _CWBCollections = New CWBCommon.CWBCollections
            End If

            Return _CWBCollections
        End Get
    End Property


    Public ReadOnly Property CWBEnums() As CWBCommon.CWBEnums
        Get

            If _CWBEnums Is Nothing Then
                _CWBEnums = New CWBCommon.CWBEnums()
            End If

            Return _CWBEnums
        End Get
    End Property
#End Region

#Region "Form Events"
    Private Sub frmReceivePayments_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.Chr(27) Then
            Me.Close()
        End If
    End Sub

    Private Sub frmReceivePayments_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetReceivePaymentsBarButton(bbSave, bbPrint, bbShowBy, bbRefresh, bbClose, bbPaid)

    End Sub

    Private Sub frmReceivePayments_Activated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        Me.PopulateUnsettleGrid()
        Me.PopulateCollectionsGrid()
    End Sub
#End Region

#Region "Bar Button Events"

    Private Sub bbClose_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbClose.ItemClick
        Me.Close()
    End Sub

    Private Sub bbRefresh_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbRefresh.ItemClick
        Select Case xTab1.SelectedTabPageIndex
            Case 0
                With gvUnsettledBills
                    .ClearColumnsFilter()
                    .ClearGrouping()
                    .ClearSelection()
                    .ClearSorting()
                End With

            Case 1
                With gvSettledBills
                    .ClearColumnsFilter()
                    .ClearGrouping()
                    .ClearSelection()
                    .ClearSorting()
                End With


        End Select

    End Sub

    Private Sub bbShowBy_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbShowBy.ItemClick
        Select Case xTab1.SelectedTabPageIndex
            Case 0
                gvUnsettledBills.Columns("CustomerNo").Group()
            Case 1
                gvSettledBills.Columns("CustomerNo").Group()

        End Select
    End Sub

    Private Sub bbSave_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbSave.ItemClick

        If Not Me.txtSalesID.Text = String.Empty Then

            Dim frm As New frmUpdateYesNo
            frm.peImage.Image = CWB.My.Resources.Resources.ImgUpdate
            frm.Text = CWB_UPDATE_CONFIRMATION_TITLE
            frm.lblTitle.Text = CWB_UPDATE_CONFIRMATION_TITLELABEL
            frm.lblDescription.Text = CWB_UPDATE_CONFIRMATION_DESCRIPTIONLABEL


            If frm.ShowDialog = Windows.Forms.DialogResult.Yes Then

                Me.SaveRecords()

            End If
        End If
    End Sub

    Private Sub bbPrint_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbPrint.ItemClick
        Try

            Select Case xTab1.SelectedTabPageIndex
                Case 0
                    PrintPreview(gcUnsettledBills, "Unsettled Sales Bills")

                Case 1
                    PrintPreview(gcSettledBills, "Settled Sales Bills")
            End Select
        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub

    Private Sub bbPaid_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbPaid.ItemClick
        Try
            If Not Me.txtSalesID.Text = String.Empty Then

                If gvUnsettledBills.GetFocusedRowCellValue(GridColumn7) = 0 Then

                    Dim frm As New frmUpdateYesNo
                    frm.peImage.Image = CWB.My.Resources.Resources.ImgUpdate
                    frm.Text = CWB_PAID_CONFIRMATION_TITLE
                    frm.lblTitle.Text = CWB_PAID_CONFIRMATION_TITLELABEL
                    frm.lblDescription.Text = CWB_PAID_CONFIRMATION_DESCRIPTIONLABEL


                    If frm.ShowDialog = Windows.Forms.DialogResult.Yes Then


                        CWBSales.SalesID = gvUnsettledBills.GetFocusedRowCellValue(GridColumn1)
                        CWBSales.UpdatedBy = UserID
                        CWBSales.SalesSetAsPaid()
                        Me.PopulateUnsettleGrid()

                    End If

                End If






            End If


        Catch ex As Exception
            MessageError(ex.ToString)
        End Try

    End Sub
#End Region

#Region "Save Records"
    Private Sub SaveRecords()
        Dim _Connection As DbConnection = Nothing
        Dim _Transaction As DbTransaction = Nothing
        Try
            Dim _DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            _Connection = _DB.CreateConnection
            _Connection.Open()
            _Transaction = _Connection.BeginTransaction()
            gvCollections.PostEditor()
            gvCollections.MoveLast()

            With CWBCollections
                .SystemID = Convert.ToInt64(IIf(txtSalesID.Text = String.Empty, 0, txtSalesID.Text))
                .TransactionTypeID = CWBCommon.CWBEnums.EnumTransactionTypes.SALES
                .CollectionDelete(_DB, _Transaction)


                'Add new Records to Collection Table
                For i As Integer = 0 To Me.gvCollections.RowCount
                    If Not gvCollections.GetRowCellDisplayText(i, gvCollections.Columns(0)) = "" Then

                        .SystemID = Convert.ToInt64(IIf(txtSalesID.Text = String.Empty, 0, txtSalesID.Text))
                        .TransactionTypeID = CWBCommon.CWBEnums.EnumTransactionTypes.SALES

                        Select Case Me.gvCollections.GetRowCellDisplayText(i, GridColumn8)

                            Case "CASH"
                                .PaymentTypeID = CWBCommon.CWBEnums.EnumPaymentTypes.CASH
                            Case "CHECK"
                                .PaymentTypeID = CWBCommon.CWBEnums.EnumPaymentTypes.CHECK
                            Case "CR CARD"
                                .PaymentTypeID = CWBCommon.CWBEnums.EnumPaymentTypes.CCARD

                        End Select

                        .Date = Date.Parse(Me.gvCollections.GetRowCellDisplayText(i, GridColumn9))
                        .Reference = Me.gvCollections.GetRowCellDisplayText(i, GridColumn10)
                        .Amount = Val(Me.gvCollections.GetRowCellDisplayText(i, GridColumn12))
                        .InsertCollection(_DB, _Transaction)
                    End If

                Next

            End With

            _Transaction.Commit()
            Dim frm As New frmSavedOk
            frm.Text = CWB_SAVESUCCESS_CONFIRMATION_TITLE
            frm.lblTitle.Text = CWB_SAVESUCCESS_CONFIRMATION_TITLELABEL
            frm.lblDescription.Text = CWB_SAVESUCCESS_CONFIRMATION_DESCRIPTIONLABEL
            frm.ShowDialog()
            Me.PopulateUnsettleGrid()


        Catch ex As Exception
            _Transaction.Rollback()
            MessageError(ex.ToString)
        Finally
            If _Connection.State = ConnectionState.Open Then
                _Connection.Close()
            End If
        End Try
    End Sub
#End Region

#Region "Populate Unsettle Sales Bills"
    Private Sub PopulateUnsettleGrid()
        Try
            gcUnsettledBills.DataSource = CWBSales.GetAllSalesUnsettledBills().Tables(0)
            Me.PopulateCollectionsGrid()
        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Populate Settled Sales Bills"
    Private Sub PopulateSettleGrid()
        Try
            gcSettledBills.DataSource = CWBSales.GetAllSalesSettledBills().Tables(0)
            Me.PopulateCollectionsGrid()

        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Populate Collections Grid"
    Private Sub PopulateCollectionsGrid()
        Try
            CWBCollections.SystemID = Convert.ToInt64(IIf(txtSalesID.Text = String.Empty, 0, txtSalesID.Text))
            CWBCollections.TransactionTypeID = CWBCommon.CWBEnums.EnumTransactionTypes.SALES
            gcCollections.DataSource = CWBCollections.CollectionGetByID().Tables(0)
        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Tab Control Events"
    Private Sub xTab1_SelectedPageChanged(ByVal sender As System.Object, ByVal e As DevExpress.XtraTab.TabPageChangedEventArgs) Handles xTab1.SelectedPageChanged
        Select Case e.Page.TabControl.SelectedTabPageIndex
            Case 0
                bbPaid.Visibility = DevExpress.XtraBars.BarItemVisibility.Always
                Me.txtSalesID.Text = String.Empty
                Me.txtCustomer.Text = String.Empty
                Me.PopulateUnsettleGrid()

            Case 1
                bbPaid.Visibility = DevExpress.XtraBars.BarItemVisibility.Never

                Me.txtSalesID.Text = String.Empty
                Me.txtCustomer.Text = String.Empty
                Me.PopulateSettleGrid()

        End Select
    End Sub
#End Region

#Region "Grid Events"
    Private Sub gcUnsettledBills_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gcUnsettledBills.Click
        Me.DisplayRecord()
    End Sub

    Private Sub gvUnsettledBills_FocusedRowChanged(ByVal sender As System.Object, ByVal e As DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs) Handles gvUnsettledBills.FocusedRowChanged
        Me.DisplayRecord()
    End Sub

    Private Sub gvCollections_CellValueChanged(ByVal sender As System.Object, ByVal e As DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs) Handles gvCollections.CellValueChanged
        Try
            Select Case e.Column.VisibleIndex
                Case 0
                    Dim GrandTotal, PaidAmount As Decimal


                    GrandTotal = Val(gvUnsettledBills.GetFocusedRowCellValue(GridColumn23))




                    GridColumn12.SummaryItem.DisplayFormat = ""
                    PaidAmount = Val(GridColumn12.SummaryText)
                    GridColumn12.SummaryItem.DisplayFormat = "{0:n2}"


                    gvCollections.SetFocusedRowCellValue(GridColumn9, Date.Today)
                    gvCollections.SetFocusedRowCellValue(GridColumn12, GrandTotal - PaidAmount)


            End Select
        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub

    Private Sub gvSettledBills_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gvSettledBills.Click
        Me.DisplayRecord()
    End Sub

    Private Sub gvSettledBills_FocusedRowChanged(ByVal sender As System.Object, ByVal e As DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs) Handles gvSettledBills.FocusedRowChanged
        Me.DisplayRecord()
    End Sub
#End Region

#Region "Display Record"
    Private Sub DisplayRecord()

        Select Case xTab1.SelectedTabPageIndex
            Case 0
                Me.txtSalesID.Text = gvUnsettledBills.GetFocusedRowCellValue(GridColumn1)
                Me.txtCustomer.Text = gvUnsettledBills.GetFocusedRowCellValue(GridColumn2)
                Me.PopulateCollectionsGrid()
            Case 1
                Me.txtSalesID.Text = gvSettledBills.GetFocusedRowCellValue(GridColumn13)
                Me.txtCustomer.Text = gvSettledBills.GetFocusedRowCellValue(GridColumn14)
                Me.PopulateCollectionsGrid()

        End Select



    End Sub
#End Region

    
    
    
End Class