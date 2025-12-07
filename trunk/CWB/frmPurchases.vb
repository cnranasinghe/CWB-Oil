Imports CWBCommon.CWBPurchases
Imports CWBCommon.CWBSuppliers
Imports CWBCommon.CWBStock
Imports Microsoft.Practices.EnterpriseLibrary.Common
Imports Microsoft.Practices.EnterpriseLibrary.Data
Imports System.Data
Imports System.Data.Common
Imports CWBCommon.CWBConstants
Imports CWBCommon.CWBCollections






Public Class frmPurchases

#Region "Variables"
    Dim _CWBPurchases As CWBCommon.CWBPurchases
    Dim _CWBSuppliers As CWBCommon.CWBSuppliers
    Dim _CWBStockItems As CWBCommon.CWBStock
    Dim _CWBCollections As CWBCommon.CWBCollections
    Dim _CWBEnums As CWBCommon.CWBEnums



#End Region

#Region "Constructor"
    Public ReadOnly Property CWBPurchases() As CWBCommon.CWBPurchases
        Get

            If _CWBPurchases Is Nothing Then
                _CWBPurchases = New CWBCommon.CWBPurchases
            End If

            Return _CWBPurchases
        End Get
    End Property

    Public ReadOnly Property CWBSuppliers() As CWBCommon.CWBSuppliers
        Get

            If _CWBSuppliers Is Nothing Then
                _CWBSuppliers = New CWBCommon.CWBSuppliers
            End If

            Return _CWBSuppliers
        End Get
    End Property

    Public ReadOnly Property CWBStockItems() As CWBCommon.CWBStock
        Get

            If _CWBStockItems Is Nothing Then
                _CWBStockItems = New CWBCommon.CWBStock
            End If

            Return _CWBStockItems
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

    Private Sub frmPurchases_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.Chr(27) Then
            Me.Close()
        End If
    End Sub

    Private Sub frmPurchasings_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetPurchaseSalesBarButton(bbSave, bbNew, bbClose, bbDisplaySelected, bbRefresh, bbPrint)
        Me.HideToolButtonsOnLoad()
        Me.dePurchaseDate.EditValue = Date.Today
        Me.deFromDate.EditValue = Date.Today
        Me.deToDate.EditValue = Date.Today
        Me.PopulateDescriptionGrid()
        Me.PopulateCollectionsGrid()

    End Sub

    Private Sub frmPurchases_Activated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        Me.PopulateSupplierLookup()
        Me.PopulateStockItemsGridLookup()
      
    End Sub
#End Region

#Region "Bar Button Events"
    Private Sub bbSave_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbSave.ItemClick
        If dxvpPurchase.Validate Then

            If lblPurchaseID.Text = String.Empty Then
                Me.SaveRecords()
                Me.PopulateStockItemsGridLookup()

            Else

                Dim frm As New frmUpdateYesNo
                frm.peImage.Image = CWB.My.Resources.Resources.ImgUpdate
                frm.Text = CWB_UPDATE_CONFIRMATION_TITLE
                frm.lblTitle.Text = CWB_UPDATE_CONFIRMATION_TITLELABEL
                frm.lblDescription.Text = CWB_UPDATE_CONFIRMATION_DESCRIPTIONLABEL

                If frm.ShowDialog = Windows.Forms.DialogResult.Yes Then
                    Me.UpdateRecords()
                    Me.PopulateStockItemsGridLookup()
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

    Private Sub bbRefresh_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbRefresh.ItemClick
        With gvPurchaseHistory
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

        PrintPreview(gcPurchaseHistory, "Purchases Report")


    End Sub
#End Region

#Region "Clear Form Data"
    Public Sub ClearFormData()

        lblPurchaseID.Text = String.Empty
        lblSystemNo.Text = String.Empty
        teReferenceNo.Text = String.Empty
        seTaxAmount.Text = "0"
        seDiscount.Text = "0"
        meNote.Text = String.Empty
        sePercentage.Text = "0"
        seGrandTotal.Text = 0
        lupSupplier.EditValue = DBNull.Value

        dxvpPurchase.RemoveControlError(lupSupplier)

        Me.PopulateDescriptionGrid()
        Me.PopulateCollectionsGrid()
        lupSupplier.Focus()

    End Sub
#End Region

#Region "Simple Button Events"
    'Private Sub SimpleButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sbCalculate.Click

    '    Try
    '        Dim a, b As Decimal

    '        GridColumn5.SummaryItem.DisplayFormat = ""
    '        a = Val(GridColumn5.SummaryText)
    '        GridColumn5.SummaryItem.DisplayFormat = "Total : {0:n2}"

    '        b = Convert.ToDecimal(a + Val(seTaxAmount.EditValue) - Val(seDiscount.EditValue))

    '        'lblBalance.Text = "Payable - " + FormatNumber(Convert.ToString(b), 2)
    '        a = 0



    '    Catch ex As Exception
    '        MessageError(ex.ToString)
    '    End Try


    'End Sub
#End Region

#Region "Populate Suppliers Lookup"
    Public Sub PopulateSupplierLookup()

        Try
            With lupSupplier
                .Properties.DataSource = CWBSuppliers.GetAllSuppliers.Tables(0)
                .Properties.DisplayMember = "SupplierName"
                .Properties.ValueMember = "SupplierID"
                .Properties.PopupWidth = 350
            End With


        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Populate Stock Items Grid Lookup"
    Public Sub PopulateStockItemsGridLookup()

        Try
            With glupStockItems
                .DataSource = CWBStockItems.GetAllStockItems.Tables(0)
                .DisplayMember = "StockCode"
                .ValueMember = "StockCode"
            End With


        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Populate Description Grid"
    Public Sub PopulateDescriptionGrid()

        Try
            With gcPurchasesDescription
                CWBPurchases.PurchaseID = Convert.ToInt64(IIf(lblPurchaseID.Text = String.Empty, 0, lblPurchaseID.Text))
                .DataSource = CWBPurchases.GetPurchaseDescriptionByID.Tables(0)

            End With


        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Populate Collections Grid"
    Private Sub PopulateCollectionsGrid()
        Try
            CWBPurchases.PurchaseID = Convert.ToInt64(IIf(lblPurchaseID.Text = String.Empty, 0, lblPurchaseID.Text))
            CWBCollections.SystemID = CWBPurchases.PurchaseID
            CWBCollections.TransactionTypeID = CWBCommon.CWBEnums.EnumTransactionTypes.PURCHASE
            gcCollections.DataSource = CWBCollections.CollectionGetByID().Tables(0)
        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Repository Events"
    Private Sub glupStockItems_EditValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles glupStockItems.EditValueChanged
        Try
            gvPurchasesDescription.SetFocusedRowCellValue(GridColumn2, glupStockItems.View.GetRowCellValue(glupStockItems.View.GetRowHandle(glupStockItems.View.GetDataSourceRowIndex(glupStockItems.GetIndexByKeyValue(gvPurchasesDescription.EditingValue))), "Description"))
            gvPurchasesDescription.SetFocusedRowCellValue(GridColumn12, glupStockItems.View.GetRowCellValue(glupStockItems.View.GetRowHandle(glupStockItems.View.GetDataSourceRowIndex(glupStockItems.GetIndexByKeyValue(gvPurchasesDescription.EditingValue))), "PurchasingPrice"))
            gvPurchasesDescription.SetFocusedRowCellValue(GridColumn17, glupStockItems.View.GetRowCellValue(glupStockItems.View.GetRowHandle(glupStockItems.View.GetDataSourceRowIndex(glupStockItems.GetIndexByKeyValue(gvPurchasesDescription.EditingValue))), "StockID"))

        Catch ex As Exception
            MessageError(ex.ToString)
        End Try

    End Sub
#End Region

#Region "Grid Events"

    Private Sub gvPurchasesDescription_CellValueChanged(ByVal sender As System.Object, ByVal e As DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs) Handles gvPurchasesDescription.CellValueChanged
        Select Case e.Column.VisibleIndex

            Case 2, 3, 4
                Dim a, b, c As Decimal

                a = IIf(Not IsDBNull(gvPurchasesDescription.GetFocusedRowCellValue(GridColumn12)), gvPurchasesDescription.GetFocusedRowCellValue(GridColumn12), 0)
                b = IIf(Not IsDBNull(gvPurchasesDescription.GetFocusedRowCellValue(GridColumn3)), gvPurchasesDescription.GetFocusedRowCellValue(GridColumn3), 0)
                c = IIf(Not IsDBNull(gvPurchasesDescription.GetFocusedRowCellValue(GridColumn4)), gvPurchasesDescription.GetFocusedRowCellValue(GridColumn4), 0)
                gvPurchasesDescription.SetFocusedRowCellValue(GridColumn5, ((a * b) - c))

                a = 0
                b = 0
                c = 0

        End Select

    End Sub

    Private Sub gvCollectionGrid_CellValueChanged(ByVal sender As System.Object, ByVal e As DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs) Handles gvCollections.CellValueChanged

        Try
            Select Case e.Column.VisibleIndex
                Case 0
                    Dim a, b As Decimal

                    GridColumn5.SummaryItem.DisplayFormat = ""
                    a = Val(GridColumn5.SummaryText)
                    GridColumn5.SummaryItem.DisplayFormat = "Total : {0:n2}"



                    GridColumn16.SummaryItem.DisplayFormat = ""
                    b = Val(GridColumn16.SummaryText)
                    GridColumn16.SummaryItem.DisplayFormat = "{0:n2}"


                    gvCollections.SetFocusedRowCellValue(GridColumn14, Date.Today)
                    gvCollections.SetFocusedRowCellValue(GridColumn16, Val(a + Val(seTaxAmount.EditValue) - Val(seDiscount.EditValue) - b))

                    Me.CalculateBalance()
            End Select
        Catch ex As Exception
            MessageError(ex.ToString)
        End Try


    End Sub

    Private Sub gcPurchasesDescription_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles gvPurchasesDescription.KeyDown
        If e.KeyCode = Keys.Delete Then
            gvPurchasesDescription.DeleteRow(gvPurchasesDescription.FocusedRowHandle)
        End If
    End Sub

    Private Sub gvCollections_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles gvCollections.KeyDown
        If e.KeyCode = Keys.Delete Then
            gvCollections.DeleteRow(gvCollections.FocusedRowHandle)
        End If
    End Sub

    Private Sub gvPurchaseHistory_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gvPurchaseHistory.DoubleClick
        Me.DisplayRecord()
        Me.CalculateBalance()
    End Sub

    Private Sub gvPurchasesDescription_RowUpdated(ByVal sender As System.Object, ByVal e As DevExpress.XtraGrid.Views.Base.RowObjectEventArgs) Handles gvPurchasesDescription.RowUpdated

        Me.CalculateBalance()
    End Sub

    Private Sub gvPurchasesDescription_FocusedRowChanged(ByVal sender As System.Object, ByVal e As DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs) Handles gvPurchasesDescription.FocusedRowChanged
        Me.CalculateBalance()
    End Sub
#End Region

#Region "Populate Purchase History"
    Private Sub PopulatePurchaseHistory()

        Try
            CWBPurchases.FromDate = Me.deFromDate.EditValue
            CWBPurchases.ToDate = Me.deToDate.EditValue
            gcPurchaseHistory.DataSource = CWBPurchases.GetPurchasesByDates().Tables(0)
        Catch ex As Exception
            MessageError(ex.ToString)
        End Try

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

    End Sub
#End Region

#Region "Show Tool Buttons On New Record Tab change"
    Public Sub ShowToolButtonsOnNewRecordTabChange()

        Me.bbDisplaySelected.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
        Me.bbRefresh.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
        Me.bbPrint.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
        Me.bbSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Always
        Me.bbNew.Visibility = DevExpress.XtraBars.BarItemVisibility.Always

    End Sub
#End Region

#Region "Tab Control Events"
    Private Sub xTab1_SelectedPageChanged(ByVal sender As System.Object, ByVal e As DevExpress.XtraTab.TabPageChangedEventArgs) Handles xTab1.SelectedPageChanged
        Select Case e.Page.TabControl.SelectedTabPageIndex
            Case 0
                Me.ShowToolButtonsOnNewRecordTabChange()
            Case 1
                Me.ShowToolButtonsOnHistoryTabChange()
                Me.PopulatePurchaseHistory()
        End Select
    End Sub
#End Region

#Region "Dispay Record"
    Public Sub DisplayRecord()
        If gvPurchaseHistory.RowCount > 0 Then
            xTab1.SelectedTabPageIndex = 0
            With CWBPurchases
                .PurchaseID = Me.gvPurchaseHistory.GetFocusedRowCellValue(GridColumn18)
                .GetPurchasesByID()
                lblPurchaseID.Text = .PurchaseID
                lblSystemNo.Text = "System No - " + .PurchaseNo.ToString
                lupSupplier.EditValue = .SupplierID
                dePurchaseDate.EditValue = .PurchaseDate
                teReferenceNo.Text = .RefBillNo
                meNote.Text = .Note
                seTaxAmount.Text = .TaxAmount
                sePercentage.Text = .TaxPercent
                seDiscount.Text = .Discount
                Me.PopulateDescriptionGrid()
                Me.PopulateCollectionsGrid()
            End With


        End If
    End Sub
#End Region

#Region "Save Record"
    Private Sub SaveRecords()



        Dim _Connection As DbConnection = Nothing
        Dim _Transaction As DbTransaction = Nothing

        Try


            Dim _DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            _Connection = _DB.CreateConnection
            _Connection.Open()
            _Transaction = _Connection.BeginTransaction()


            gvPurchasesDescription.PostEditor()
            gvCollections.PostEditor()

            gvPurchasesDescription.MoveLast()
            gvCollections.MoveLast()



            With CWBPurchases
                .PurchaseID = Convert.ToInt64(IIf(lblPurchaseID.Text = String.Empty, 0, lblPurchaseID.Text))
                .SupplierID = lupSupplier.EditValue
                .PurchaseDate = dePurchaseDate.Text
                .RefBillNo = teReferenceNo.Text.Trim

                GridColumn5.SummaryItem.DisplayFormat = ""
                .Total = gvPurchasesDescription.Columns("Value").SummaryText
                .TaxPercent = sePercentage.EditValue

                .TaxAmount = seTaxAmount.EditValue
                .Discount = seDiscount.EditValue
                .GrandTotal = seGrandTotal.EditValue
                .Note = meNote.Text

                If seGrandTotal.EditValue = GridColumn16.SummaryText Then
                    .Paid = True
                Else
                    .Paid = False
                End If


                .CreatedBy = UserID
                .UpdatedBy = UserID
                .InsertPurchase(_DB, _Transaction)



                For i As Integer = 0 To Me.gvPurchasesDescription.RowCount
                    If Not gvPurchasesDescription.GetRowCellDisplayText(i, gvPurchasesDescription.Columns(0)) = "" Then
                        .PurchaseID = .CurrentPurchaseID
                        .StockID = Me.gvPurchasesDescription.GetRowCellDisplayText(i, GridColumn17)
                        .Unit_Price = Val(Me.gvPurchasesDescription.GetRowCellDisplayText(i, GridColumn12))
                        .Quantity = Val(Me.gvPurchasesDescription.GetRowCellDisplayText(i, GridColumn3))
                        .Discount = Val(Me.gvPurchasesDescription.GetRowCellDisplayText(i, GridColumn4))
                        .Value = Val(Me.gvPurchasesDescription.GetRowCellDisplayText(i, GridColumn5))
                        .InsertPurchaseDescription(_DB, _Transaction)

                        'Update Main Stock
                        .StockID = Me.gvPurchasesDescription.GetRowCellDisplayText(i, GridColumn17)
                        .StockBalance = Val(Me.gvPurchasesDescription.GetRowCellDisplayText(i, GridColumn3))
                        .UpdateStock(_DB, _Transaction)


                    End If

                Next



                With CWBCollections
                    For i As Integer = 0 To Me.gvCollections.RowCount
                        If Not gvCollections.GetRowCellDisplayText(i, gvCollections.Columns(0)) = "" Then
                            .SystemID = CWBPurchases.CurrentPurchaseID
                            .TransactionTypeID = CWBCommon.CWBEnums.EnumTransactionTypes.PURCHASE

                            Select Case Me.gvCollections.GetRowCellDisplayText(i, GridColumn13)

                                Case "CASH"
                                    .PaymentTypeID = CWBCommon.CWBEnums.EnumPaymentTypes.CASH
                                Case "CHECK"
                                    .PaymentTypeID = CWBCommon.CWBEnums.EnumPaymentTypes.CHECK
                                Case "CR CARD"
                                    .PaymentTypeID = CWBCommon.CWBEnums.EnumPaymentTypes.CCARD

                            End Select

                            .Date = Date.Parse(Me.gvCollections.GetRowCellDisplayText(i, gvCollections.Columns(1)))
                            .Reference = Me.gvCollections.GetRowCellDisplayText(i, gvCollections.Columns(2))
                            .Amount = FormatNumber(Me.gvCollections.GetRowCellDisplayText(i, gvCollections.Columns(3)), 2, , , )
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
                Me.ClearFormData()


            End With


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

#Region "Update Records"
    Private Sub UpdateRecords()


        Dim _Connection As DbConnection = Nothing
        Dim _Transaction As DbTransaction = Nothing

        Try
            Dim _DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            _Connection = _DB.CreateConnection
            _Connection.Open()
            _Transaction = _Connection.BeginTransaction()


            gvPurchasesDescription.PostEditor()
            gvCollections.PostEditor()

            gvPurchasesDescription.MoveLast()
            gvCollections.MoveLast()



            With CWBPurchases
                .PurchaseID = Convert.ToInt64(IIf(lblPurchaseID.Text = String.Empty, 0, lblPurchaseID.Text))
                .SupplierID = lupSupplier.EditValue
                .PurchaseDate = dePurchaseDate.Text
                .RefBillNo = teReferenceNo.Text.Trim

                GridColumn5.SummaryItem.DisplayFormat = ""
                .Total = gvPurchasesDescription.Columns("Value").SummaryText
                .TaxPercent = 0
                .TaxAmount = seTaxAmount.EditValue
                .Discount = seDiscount.EditValue
                .GrandTotal = seGrandTotal.EditValue
                .TaxPercent = sePercentage.EditValue
                .Note = meNote.Text
                If seGrandTotal.EditValue = GridColumn16.SummaryText Then
                    .Paid = True
                Else
                    .Paid = False
                End If

                .CreatedBy = UserID
                .UpdatedBy = UserID
                .InsertPurchase(_DB, _Transaction)

                .RemoveFromStock(_DB, _Transaction)
                .PurchasesDescriptionDelete(_DB, _Transaction)


                For i As Integer = 0 To Me.gvPurchasesDescription.RowCount
                    If Not gvPurchasesDescription.GetRowCellDisplayText(i, gvPurchasesDescription.Columns(0)) = "" Then

                        .StockID = Me.gvPurchasesDescription.GetRowCellDisplayText(i, GridColumn17)
                        .Unit_Price = Val(Me.gvPurchasesDescription.GetRowCellDisplayText(i, GridColumn12))
                        .Quantity = Val(Me.gvPurchasesDescription.GetRowCellDisplayText(i, GridColumn3))
                        .Discount = Val(Me.gvPurchasesDescription.GetRowCellDisplayText(i, GridColumn4))
                        .Value = Val(Me.gvPurchasesDescription.GetRowCellDisplayText(i, GridColumn5))
                        .InsertPurchaseDescription(_DB, _Transaction)

                        'Update Main Stock
                        .StockID = Me.gvPurchasesDescription.GetRowCellDisplayText(i, GridColumn17)
                        .StockBalance = Val(Me.gvPurchasesDescription.GetRowCellDisplayText(i, GridColumn3))
                        .UpdateStock(_DB, _Transaction)


                    End If

                Next



                With CWBCollections
                    .SystemID = Convert.ToInt64(IIf(lblPurchaseID.Text = String.Empty, 0, lblPurchaseID.Text))
                    .TransactionTypeID = CWBCommon.CWBEnums.EnumTransactionTypes.PURCHASE
                    .CollectionDelete(_DB, _Transaction)


                    'Add new Records to Collection Table
                    For i As Integer = 0 To Me.gvCollections.RowCount
                        If Not gvCollections.GetRowCellDisplayText(i, gvCollections.Columns(0)) = "" Then
                            .SystemID = Convert.ToInt64(IIf(lblPurchaseID.Text = String.Empty, 0, lblPurchaseID.Text))
                            .TransactionTypeID = CWBCommon.CWBEnums.EnumTransactionTypes.PURCHASE

                            Select Case Me.gvCollections.GetRowCellDisplayText(i, GridColumn13)

                                Case "CASH"
                                    .PaymentTypeID = CWBCommon.CWBEnums.EnumPaymentTypes.CASH
                                Case "CHECK"
                                    .PaymentTypeID = CWBCommon.CWBEnums.EnumPaymentTypes.CHECK
                                Case "CR CARD"
                                    .PaymentTypeID = CWBCommon.CWBEnums.EnumPaymentTypes.CCARD

                            End Select

                            .Date = Date.Parse(Me.gvCollections.GetRowCellDisplayText(i, gvCollections.Columns(1)))
                            .Reference = Me.gvCollections.GetRowCellDisplayText(i, gvCollections.Columns(2))
                            .Amount = FormatNumber(Me.gvCollections.GetRowCellDisplayText(i, gvCollections.Columns(3)), 2, , , )
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
                Me.ClearFormData()


            End With

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

#Region "Show Print Preview"
    Private Sub ShowPrintPreview()
        Try
            Dim _frmPrintPreview As New frmPrint
            ''Dim _xrPurchasing As New xrPurchasing
            ''_xrPurchasing.Purchases_GetByIDTableAdapter.GetData(Me.gvPurchaseHistory.GetFocusedRowCellValue(GridColumn18))
            ''_xrPurchasing.PurchasesDescription_GetByIDTableAdapter.GetData(Me.gvPurchaseHistory.GetFocusedRowCellValue(GridColumn18))
            ''_frmPrintPreview.PrintControl1.PrintingSystem = _xrPurchasing.PrintingSystem

            '_xrPurchasing.CreateDocument()

            _frmPrintPreview.MdiParent = frmMain
            _frmPrintPreview.Show()
            _frmPrintPreview.BringToFront()
        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Calculate Balance"
    Private Sub CalculateBalance()
        Try

            Dim total, taxamount, discount, balance As Decimal



            GridColumn5.SummaryItem.DisplayFormat = ""
            total = Val(GridColumn5.SummaryText)
            GridColumn5.SummaryItem.DisplayFormat = "Total : {0:n2}"


            taxamount = Me.seTaxAmount.EditValue
            discount = Me.seDiscount.EditValue
            balance = total + taxamount - (discount)

            seGrandTotal.EditValue = balance

        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Calculate Discount Percent"
    Private Sub CalculateDiscountPercent()

        Try

            Dim total, discount As Decimal



            GridColumn5.SummaryItem.DisplayFormat = ""
            total = Val(GridColumn5.SummaryText)
            GridColumn5.SummaryItem.DisplayFormat = "Total : {0:n2}"



            discount = total * (sePercentage.EditValue / 100)


            seDiscount.EditValue = discount



        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
#End Region

#Region "SpinEdit Events"

    Private Sub seTaxAmount_EditValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles seTaxAmount.EditValueChanged
        Me.CalculateBalance()
    End Sub

    Private Sub seDiscount_EditValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles seDiscount.EditValueChanged
        Me.CalculateBalance()
    End Sub

    Private Sub sePercentage_EditValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sePercentage.EditValueChanged
        Me.CalculateDiscountPercent()
    End Sub
#End Region

#Region "Button Events"
    Private Sub smProcess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles smProcess.Click
        If dxvpHistoryData.Validate Then
            Me.PopulatePurchaseHistory()
        End If

    End Sub
#End Region


End Class