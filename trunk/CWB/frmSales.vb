Imports CWBCommon.CWBEmployers
Imports CWBCommon.CWBSales
Imports CWBCommon.CWBCustomers
Imports CWBCommon.CWBStock
Imports Microsoft.Practices.EnterpriseLibrary.Common
Imports Microsoft.Practices.EnterpriseLibrary.Data
Imports System.Data
Imports System.Data.Common
Imports CWBCommon.CWBConstants
Imports CWBCommon.CWBModels





Public Class frmSales

#Region "Variables"
    Dim _CWBEmployers As CWBCommon.CWBEmployers
    Dim _CWBSales As CWBCommon.CWBSales
    Dim _CWBCustomers As CWBCommon.CWBCustomers
    Dim _CWBStockItems As CWBCommon.CWBStock
    Dim _CWBCollections As CWBCommon.CWBCollections
    Dim _CWBEnums As CWBCommon.CWBEnums
    Dim _CWBModels As CWBCommon.CWBModels
    Dim _CWBApplication As CWBCommon.CWBApplication
    Dim _CWBMessages As CWBCommon.CWBMessages



#End Region

#Region "Constructor"

    Public ReadOnly Property CWBEmployers() As CWBCommon.CWBEmployers
        Get

            If _CWBEmployers Is Nothing Then
                _CWBEmployers = New CWBCommon.CWBEmployers
            End If

            Return _CWBEmployers
        End Get
    End Property

    Public ReadOnly Property CWBModels() As CWBCommon.CWBModels
        Get

            If _CWBModels Is Nothing Then
                _CWBModels = New CWBCommon.CWBModels
            End If

            Return _CWBModels
        End Get
    End Property

    Public ReadOnly Property CWBSales() As CWBCommon.CWBSales
        Get

            If _CWBSales Is Nothing Then
                _CWBSales = New CWBCommon.CWBSales
            End If

            Return _CWBSales
        End Get
    End Property

    Public ReadOnly Property CWBCustomers() As CWBCommon.CWBCustomers
        Get

            If _CWBCustomers Is Nothing Then
                _CWBCustomers = New CWBCommon.CWBCustomers
            End If

            Return _CWBCustomers
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
    Public ReadOnly Property CWBStockItems() As CWBCommon.CWBStock
        Get

            If _CWBStockItems Is Nothing Then
                _CWBStockItems = New CWBCommon.CWBStock
            End If

            Return _CWBStockItems
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

    Public ReadOnly Property CWBApplication() As CWBCommon.CWBApplication
        Get

            If _CWBApplication Is Nothing Then
                _CWBApplication = New CWBCommon.CWBApplication
            End If

            Return _CWBApplication
        End Get
    End Property


    Public ReadOnly Property CWBMessages() As CWBCommon.CWBMessages
        Get

            If _CWBMessages Is Nothing Then
                _CWBMessages = New CWBCommon.CWBMessages
            End If

            Return _CWBMessages
        End Get
    End Property
#End Region

#Region "Form Events"

    Private Sub frmSales_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.Chr(27) Then
            Me.Close()
        End If
    End Sub

    Private Sub frmSales_Activated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        Me.PopulateCustomerLookup()
        Me.PopulateStockItemsGridLookup()
        Me.PopulateEmplpoyers()
    End Sub

    Private Sub frmSales_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetPurchaseSalesBarButton(bbSave, bbNew, bbClose, bbDisplaySelected, bbRefresh, bbPrint, bbAudit)
        Me.HideToolButtonsOnLoad()
        Me.PopulateModelLookup()
        Me.PopulateDescriptionGrid()
        Me.PopulateCollectionsGrid()
        Me.PopulateMessageLookup()

        Me.deSalesDate.EditValue = Date.Today

        If UserType = "USER" Then

            Me.deSalesDate.Enabled = False
        End If

        Me.deFromDate.EditValue = Date.Today
        Me.deToDate.EditValue = Date.Today
    End Sub

  


#End Region

#Region "Bar Button Events"
    Private Sub bbSave_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbSave.ItemClick
        If dxvpSales.Validate Then

            If lblSalesID.Text = String.Empty Then
                Me.SaveRecords()
                Me.PopulateStockItemsGridLookup()

            Else

                Dim frm As New frmUpdateYesNo
                frm.peImage.Image = CWB.My.Resources.Resources.ImgUpdate
                frm.Text = CWB_UPDATE_CONFIRMATION_TITLE
                frm.lblTitle.Text = CWB_UPDATE_CONFIRMATION_TITLELABEL
                frm.lblDescription.Text = CWB_UPDATE_CONFIRMATION_DESCRIPTIONLABEL

                If UserType = "OWNER" Then
                    If frm.ShowDialog = Windows.Forms.DialogResult.Yes Then
                        Me.UpdateRecords()
                        Me.PopulateStockItemsGridLookup()
                    End If

                ElseIf UserType = "USER" Then

                    If deSalesDate.EditValue = CWBApplication.GetServerDateTime Then
                        If frm.ShowDialog = Windows.Forms.DialogResult.Yes Then
                            Me.UpdateRecords()
                            Me.PopulateStockItemsGridLookup()
                        End If
                    Else
                        MessageOK("Can not Update", "You don't have permission", "Click OK to continue")

                    End If

                   
                Else
                    MessageOK("Can not Update", "You don't have permission", "Click OK to continue")
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
        With gvSalesHistory
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

        Select Case xTab1.SelectedTabPageIndex
            Case 0
                If Not Me.lblSalesID.Text = String.Empty Then

                    Me.ShowPrintPreview(Convert.ToInt64(Me.lblSalesID.Text))
                End If
            Case 1
                If gvSalesHistory.RowCount > 0 Then

                    PrintPreview(gcSalesHistory, "Sales Report")
                Else
                    Exit Sub
                End If

        End Select


    End Sub

#End Region

#Region "Clear Form Data"
    Public Sub ClearFormData()

        lblSalesID.Text = String.Empty
        teReferenceNo.Text = String.Empty
        lblSystemNo.Text = String.Empty
        seTaxAmount.Text = "0"
        seDiscount.Text = "0"
        sbCalculate.Text = "Calculate"
        seGrandTotal.EditValue = 0
        sePercentage.Text = "0"
        lupCustomer.EditValue = DBNull.Value
        seCurrentMeterReading.Text = "0"
        seNextService.Text = "0"
        leMessages.EditValue = Nothing
        leService1.EditValue = Nothing
        leService2.EditValue = Nothing
        leModal.EditValue = Nothing
        clbCh.Items(0).CheckState = CheckState.Unchecked
        clbCh.Items(1).CheckState = CheckState.Unchecked
        clbCh.Items(2).CheckState = CheckState.Unchecked
        clbCh.Items(3).CheckState = CheckState.Unchecked

        lcPreviousBills.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never
        dxvpSales.RemoveControlError(lupCustomer)
        dxvpSales.RemoveControlError(leModal)
        Me.PopulateDescriptionGrid()
        Me.PopulateCollectionsGrid()
        lupCustomer.Focus()

    End Sub
#End Region

#Region "Simple Button Events"
    Private Sub SimpleButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sbCalculate.Click

        Try
            Dim a As Decimal

            GridColumn5.SummaryItem.DisplayFormat = ""
            a = Val(GridColumn5.SummaryText)
            GridColumn5.SummaryItem.DisplayFormat = "Total : {0:n2}"


            sbCalculate.Text = a + Val(seTaxAmount.EditValue) - Val(seDiscount.EditValue)
            a = 0

            gvCollections.Focus()
            gvCollections.AddNewRow()
        Catch ex As Exception
            MessageError(ex.ToString)
        End Try


    End Sub
#End Region

#Region "Populate Customer Lookup"
    Public Sub PopulateCustomerLookup()

        Try
            With lupCustomer
                .Properties.DataSource = CWBCustomers.GetAllCustomers().Tables(0)
                .Properties.DisplayMember = "CustomerName"
                .Properties.ValueMember = "CustomerID"
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
            With gcSalesDescription
                CWBSales.SalesID = Convert.ToInt64(IIf(lblSalesID.Text = String.Empty, 0, lblSalesID.Text))
                .DataSource = CWBSales.GetSalesDescriptionByID.Tables(0)

            End With


        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Populate Collections Grid"
    Private Sub PopulateCollectionsGrid()
        Try
            CWBSales.SalesID = Convert.ToInt64(IIf(lblSalesID.Text = String.Empty, 0, lblSalesID.Text))
            CWBCollections.SystemID = CWBSales.SalesID
            CWBCollections.TransactionTypeID = CWBCommon.CWBEnums.EnumTransactionTypes.SALES
            gcCollections.DataSource = CWBCollections.CollectionGetByID().Tables(0)
        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Populate  Employers"
    Private Sub PopulateEmplpoyers()
        Try

            leService1.Properties.DataSource = CWBEmployers.SelectAll.Tables(0)
            leService1.Properties.DisplayMember = "EmployerName"
            leService1.Properties.ValueMember = "EmployerID"


            leService2.Properties.DataSource = CWBEmployers.SelectAll.Tables(0)
            leService2.Properties.DisplayMember = "EmployerName"
            leService2.Properties.ValueMember = "EmployerID"

        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Repository Events"
    Private Sub glupStockItems_EditValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles glupStockItems.EditValueChanged
        Try
            'gvSalesDescription.SetFocusedRowCellValue(GridColumn36, glupStockItems.View.GetRowCellValue(glupStockItems.View.GetRowHandle(glupStockItems.View.GetDataSourceRowIndex(glupStockItems.GetIndexByKeyValue(gvSalesDescription.EditingValue))), "Description"))
            'gvSalesDescription.SetFocusedRowCellValue(GridColumn37, glupStockItems.View.GetRowCellValue(glupStockItems.View.GetRowHandle(glupStockItems.View.GetDataSourceRowIndex(glupStockItems.GetIndexByKeyValue(gvSalesDescription.EditingValue))), "SalesPrice"))
            'gvSalesDescription.SetFocusedRowCellValue(GridColumn28, glupStockItems.View.GetRowCellValue(glupStockItems.View.GetRowHandle(glupStockItems.View.GetDataSourceRowIndex(glupStockItems.GetIndexByKeyValue(gvSalesDescription.EditingValue))), "StockID"))
            'gvSalesDescription.SetFocusedRowCellValue(GridColumn47, glupStockItems.View.GetRowCellValue(glupStockItems.View.GetRowHandle(glupStockItems.View.GetDataSourceRowIndex(glupStockItems.GetIndexByKeyValue(gvSalesDescription.EditingValue))), "PurchasingPrice"))


            Dim gridLookUpEdit As DevExpress.XtraEditors.GridLookUpEdit = CType(sender, DevExpress.XtraEditors.GridLookUpEdit)
            Dim iCurrentRow As Integer = CInt(gvPurchasesDescription.FocusedRowHandle.ToString)

            If gridLookUpEdit.Properties.View.GetFocusedRowCellValue(gcolStockCode) IsNot Nothing Then
                gvSalesDescription.SetFocusedRowCellValue(colStockID, gridLookUpEdit.Properties.View.GetFocusedRowCellValue(gcolStockId).ToString())
                gvSalesDescription.SetFocusedRowCellValue(colDescription, gridLookUpEdit.Properties.View.GetFocusedRowCellValue(gcolDescription).ToString())
                gvSalesDescription.SetFocusedRowCellValue(colUnitPrice, gridLookUpEdit.Properties.View.GetFocusedRowCellValue(gcolSalesPrice).ToString())
                gvSalesDescription.SetFocusedRowCellValue(colPurchasePrice, gridLookUpEdit.Properties.View.GetFocusedRowCellValue(gcolPurchasePrice).ToString())

            End If


        Catch ex As Exception
            MessageError(ex.ToString)
        End Try


    End Sub
#End Region

#Region "Grid Events"

    Private Sub gvSalesDescription_CellValueChanged(ByVal sender As System.Object, ByVal e As DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs) Handles gvSalesDescription.CellValueChanged

        Try
            Select Case e.Column.VisibleIndex

                Case 2, 3, 4
                    Dim a, b, c As Decimal

                    a = IIf(Not IsDBNull(gvSalesDescription.GetFocusedRowCellValue(GridColumn12)), gvSalesDescription.GetFocusedRowCellValue(GridColumn12), 0)
                    b = IIf(Not IsDBNull(gvSalesDescription.GetFocusedRowCellValue(GridColumn3)), gvSalesDescription.GetFocusedRowCellValue(GridColumn3), 0)
                    c = IIf(Not IsDBNull(gvSalesDescription.GetFocusedRowCellValue(GridColumn4)), gvSalesDescription.GetFocusedRowCellValue(GridColumn4), 0)
                    gvSalesDescription.SetFocusedRowCellValue(GridColumn5, ((a * b) - c))
                    a = 0
                    b = 0
                    c = 0
            End Select
        Catch ex As Exception
            MessageError(ex.ToString)
        End Try


    End Sub


    Private Sub gcSalessDescription_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles gvSalesDescription.KeyDown
        If e.KeyCode = Keys.Delete Then
            gvSalesDescription.DeleteRow(gvSalesDescription.FocusedRowHandle)
        End If
    End Sub

    Private Sub gcCollections_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles gcCollections.KeyDown, gvCollections.KeyDown
        If e.KeyCode = Keys.Delete Then
            gvCollections.DeleteRow(gvCollections.FocusedRowHandle)
        End If
    End Sub

    Private Sub gvSalesHistory_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gvSalesHistory.DoubleClick
        Me.DisplayRecord()
    End Sub

    Private Sub gvSalesDescription_FocusedRowChanged(ByVal sender As System.Object, ByVal e As DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs) Handles gvSalesDescription.FocusedRowChanged
        Me.CalculateBalance()
    End Sub

    Private Sub gvSalesDescription_RowUpdated(ByVal sender As System.Object, ByVal e As DevExpress.XtraGrid.Views.Base.RowObjectEventArgs) Handles gvSalesDescription.RowUpdated
        Me.CalculateBalance()
    End Sub

    Private Sub gvCollections_CellValueChanged(ByVal sender As System.Object, ByVal e As DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs) Handles gvCollections.CellValueChanged
        Try

            Select Case e.Column.VisibleIndex
                Case 0
                    Dim a, b As Decimal

                    GridColumn40.SummaryItem.DisplayFormat = ""
                    a = Val(GridColumn40.SummaryText)
                    GridColumn40.SummaryItem.DisplayFormat = "Total : {0:n2}"

                    GridColumn27.SummaryItem.DisplayFormat = ""
                    b = Val(GridColumn27.SummaryText)
                    GridColumn27.SummaryItem.DisplayFormat = "{0:n2}"

                    gvCollections.SetFocusedRowCellValue(GridColumn25, Date.Today)
                    gvCollections.SetFocusedRowCellValue(GridColumn27, Val(a + Val(seTaxAmount.EditValue) - Val(seDiscount.EditValue) - b))


            End Select
        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Populate Sales History"
    Private Sub PopulateSalesHistory()

        Try


            CWBSales.FromDate = Me.deFromDate.EditValue
            CWBSales.ToDate = Me.deToDate.EditValue
            gcSalesHistory.DataSource = CWBSales.GetSalesByDates().Tables(0)
        Catch ex As Exception
            MessageError(ex.ToString)
        End Try

    End Sub
#End Region

#Region "Hide Tool Buttons On Load"
    Public Sub HideToolButtonsOnLoad()

        Me.bbDisplaySelected.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
        Me.bbRefresh.Visibility = DevExpress.XtraBars.BarItemVisibility.Never


    End Sub
#End Region

#Region "Show Tool Buttons On History Tab change"
    Public Sub ShowToolButtonsOnHistoryTabChange()

        Me.bbDisplaySelected.Visibility = DevExpress.XtraBars.BarItemVisibility.Always
        Me.bbRefresh.Visibility = DevExpress.XtraBars.BarItemVisibility.Always
        Me.bbPrint.Visibility = DevExpress.XtraBars.BarItemVisibility.Always
        Me.bbSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
        Me.bbNew.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
        Me.bbAudit.Visibility = DevExpress.XtraBars.BarItemVisibility.Never

    End Sub
#End Region

#Region "Show Tool Buttons On New Record Tab change"
    Public Sub ShowToolButtonsOnNewRecordTabChange()

        Me.bbDisplaySelected.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
        Me.bbRefresh.Visibility = DevExpress.XtraBars.BarItemVisibility.Never

        Me.bbSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Always
        Me.bbNew.Visibility = DevExpress.XtraBars.BarItemVisibility.Always
        Me.bbAudit.Visibility = DevExpress.XtraBars.BarItemVisibility.Always
    End Sub
#End Region

#Region "Tab Control Events"
    Private Sub xTab1_SelectedPageChanged(ByVal sender As System.Object, ByVal e As DevExpress.XtraTab.TabPageChangedEventArgs) Handles xTab1.SelectedPageChanged
        Select Case e.Page.TabControl.SelectedTabPageIndex
            Case 0
                Me.ShowToolButtonsOnNewRecordTabChange()
            Case 1
                Me.ShowToolButtonsOnHistoryTabChange()
                Me.PopulateSalesHistory()

        End Select
    End Sub
#End Region

#Region "Dispay Record"
    Public Sub DisplayRecord()

        Try
            If gvSalesHistory.RowCount > 0 Then
                xTab1.SelectedTabPageIndex = 0
                With CWBSales
                    .SalesID = Me.gvSalesHistory.GetFocusedRowCellValue(GridColumn41)
                    .GetSalesByID()
                    lblSalesID.Text = .SalesID
                    lblSystemNo.Text = "System No - " + .SalesNo.ToString

                    lupCustomer.EditValue = .CustomerID
                    deSalesDate.EditValue = .SalesDate
                    teReferenceNo.Text = .ReferenceNo
                    leModal.EditValue = .ModelID
                    seCurrentMeterReading.EditValue = .CurrenMeterReading
                    seNextService.EditValue = .NextMeterReading
                    leService1.EditValue = .Service1
                    leService2.EditValue = .Service2
                    sePercentage.EditValue = .TaxPercent

                    If .Ch1 = True Then
                        clbCh.Items(0).CheckState = CheckState.Checked
                    Else
                        clbCh.Items(0).CheckState = CheckState.Unchecked
                    End If


                    If .Ch2 = True Then
                        clbCh.Items(1).CheckState = CheckState.Checked
                    Else
                        clbCh.Items(1).CheckState = CheckState.Unchecked
                    End If


                    If .Ch3 = True Then
                        clbCh.Items(2).CheckState = CheckState.Checked
                    Else
                        clbCh.Items(2).CheckState = CheckState.Unchecked
                    End If


                    If .Ch4 = True Then
                        clbCh.Items(3).CheckState = CheckState.Checked
                    Else
                        clbCh.Items(3).CheckState = CheckState.Unchecked
                    End If
                    leMessages.EditValue = .MessageID
                    sePercentage.Text = .TaxPercent
                    'meNote.Text = .Note
                    seTaxAmount.Text = .TaxAmount
                    seDiscount.Text = .Discount
                    Me.PopulateDescriptionGrid()
                    Me.PopulateCollectionsGrid()
                End With


            End If
        Catch ex As Exception
            MessageError(ex.ToString)
        End Try


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


            gvSalesDescription.PostEditor()
            gvCollections.PostEditor()

            gvSalesDescription.MoveLast()
            gvCollections.MoveLast()


            With CWBSales
                .SalesID = Convert.ToInt64(IIf(lblSalesID.Text = String.Empty, 0, lblSalesID.Text))
                .CustomerID = lupCustomer.EditValue

                If UserType = "OWNER" Then
                    .SalesDate = deSalesDate.Text
                Else
                    .SalesDate = CWBApplication.GetServerDateTime
                End If

                .ReferenceNo = teReferenceNo.Text.Trim
                .ModelID = leModal.EditValue
                .CurrenMeterReading = seCurrentMeterReading.Text
                .NextMeterReading = seNextService.Text

                GridColumn40.SummaryItem.DisplayFormat = ""
                .Total = gvSalesDescription.Columns("Value").SummaryText
                .TaxPercent = sePercentage.Text
                .TaxAmount = 0 'seTaxAmount.EditValue
                .Discount = seDiscount.EditValue
                .GrandTotal = seGrandTotal.EditValue

                .Service1 = leService1.EditValue
                .Service2 = leService2.EditValue
                .MessageID = leMessages.EditValue

                If clbCh.Items(0).CheckState = CheckState.Checked Then
                    .Ch1 = True

                Else
                    .Ch1 = False
                End If

                If clbCh.Items(1).CheckState = CheckState.Checked Then
                    .Ch2 = True

                Else
                    .Ch2 = False
                End If


                If clbCh.Items(2).CheckState = CheckState.Checked Then
                    .Ch3 = True

                Else
                    .Ch3 = False
                End If

                If clbCh.Items(3).CheckState = CheckState.Checked Then
                    .Ch4 = True

                Else
                    .Ch4 = False
                End If

                If seGrandTotal.EditValue = GridColumn27.SummaryText Then
                    .Paid = True
                Else
                    .Paid = False
                End If



                .CreatedBy = UserID
                .UpdatedBy = UserID
                .InsertSales(_DB, _Transaction)


                For i As Integer = 0 To Me.gvSalesDescription.RowCount
                    If Not gvSalesDescription.GetRowCellDisplayText(i, gvSalesDescription.Columns(0)) = "" Then
                        .SalesID = .CurrentSalesID
                        .StockID = Me.gvSalesDescription.GetRowCellDisplayText(i, colStockID)
                        .UnitPrice = Val(Me.gvSalesDescription.GetRowCellDisplayText(i, colUnitPrice))
                        .PurchasingPrice = Val(Me.gvSalesDescription.GetRowCellDisplayText(i, colPurchasePrice))
                        .Quantity = Val(Me.gvSalesDescription.GetRowCellDisplayText(i, colQuantity))
                        .Discount = Val(Me.gvSalesDescription.GetRowCellDisplayText(i, colDiscountAmt))

                        .Value = Val(Me.gvSalesDescription.GetRowCellDisplayText(i, GridColumn40))
                        .InsertSalesDescription(_DB, _Transaction)

                        'Update Main Stock
                        .StockID = Me.gvSalesDescription.GetRowCellDisplayText(i, colStockID)
                        .StockBalance = Val(Me.gvSalesDescription.GetRowCellDisplayText(i, colQuantity))
                        .UpdateStock(_DB, _Transaction)


                    End If

                Next



                With CWBCollections
                    For i As Integer = 0 To Me.gvCollections.RowCount
                        If Not gvCollections.GetRowCellDisplayText(i, gvCollections.Columns(0)) = "" Then
                            .SystemID = CWBSales.CurrentSalesID
                            .TransactionTypeID = CWBCommon.CWBEnums.EnumTransactionTypes.SALES

                            Select Case Me.gvCollections.GetRowCellDisplayText(i, GridColumn22)

                                Case "CASH"
                                    .PaymentTypeID = CWBCommon.CWBEnums.EnumPaymentTypes.CASH
                                Case "CHECK"
                                    .PaymentTypeID = CWBCommon.CWBEnums.EnumPaymentTypes.CHECK
                                Case "CR CARD"
                                    .PaymentTypeID = CWBCommon.CWBEnums.EnumPaymentTypes.CCARD

                            End Select

                            .Date = Date.Parse(Me.gvCollections.GetRowCellDisplayText(i, GridColumn25))
                            .Reference = Me.gvCollections.GetRowCellDisplayText(i, GridColumn26)

                            .Amount = FormatNumber(Me.gvCollections.GetRowCellDisplayText(i, GridColumn27), 2, , , )
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


                ShowPrintPreview(CWBSales.CurrentSalesID)

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


            gvSalesDescription.PostEditor()
            gvCollections.PostEditor()

            gvSalesDescription.MoveLast()
            gvCollections.MoveLast()



            With CWBSales
                .SalesID = Convert.ToInt64(IIf(lblSalesID.Text = String.Empty, 0, lblSalesID.Text))
                .CustomerID = lupCustomer.EditValue
                .SalesDate = deSalesDate.Text
                .ReferenceNo = teReferenceNo.Text.Trim
                .ModelID = leModal.EditValue
                .CurrenMeterReading = seCurrentMeterReading.Text
                .NextMeterReading = seNextService.Text
                .TaxPercent = sePercentage.Text


                GridColumn40.SummaryItem.DisplayFormat = ""
                .Total = gvSalesDescription.Columns("Value").SummaryText


                .TaxAmount = 0 ' seTaxAmount.EditValue
                .Discount = seDiscount.EditValue
                .GrandTotal = seGrandTotal.EditValue

                .Service1 = leService1.EditValue
                .Service2 = leService2.EditValue
                .MessageID = leMessages.EditValue

                If clbCh.Items(0).CheckState = CheckState.Checked Then
                    .Ch1 = True

                Else
                    .Ch1 = False
                End If

                If clbCh.Items(1).CheckState = CheckState.Checked Then
                    .Ch2 = True

                Else
                    .Ch2 = False
                End If


                If clbCh.Items(2).CheckState = CheckState.Checked Then
                    .Ch3 = True

                Else
                    .Ch3 = False
                End If

                If clbCh.Items(3).CheckState = CheckState.Checked Then
                    .Ch4 = True

                Else
                    .Ch4 = False
                End If

                If seGrandTotal.EditValue = GridColumn27.SummaryText Then
                    .Paid = True
                Else
                    .Paid = False
                End If

                .CreatedBy = UserID
                .UpdatedBy = UserID
                .InsertSales(_DB, _Transaction)

                .AddToStock(_DB, _Transaction)
                .SalesDescriptionDelete(_DB, _Transaction)


                For i As Integer = 0 To Me.gvSalesDescription.RowCount
                    If Not gvSalesDescription.GetRowCellDisplayText(i, gvSalesDescription.Columns(0)) = "" Then

                        .StockID = Me.gvSalesDescription.GetRowCellDisplayText(i, colStockID)
                        .UnitPrice = Val(Me.gvSalesDescription.GetRowCellDisplayText(i, colUnitPrice))
                        .PurchasingPrice = Val(Me.gvSalesDescription.GetRowCellDisplayText(i, colPurchasePrice))
                        .Quantity = Val(Me.gvSalesDescription.GetRowCellDisplayText(i, colQuantity))
                        .Discount = Val(Me.gvSalesDescription.GetRowCellDisplayText(i, colDiscountAmt))

                        .Value = Val(Me.gvSalesDescription.GetRowCellDisplayText(i, GridColumn40))
                        .InsertSalesDescription(_DB, _Transaction)

                        'Update Main Stock
                        .StockID = Me.gvSalesDescription.GetRowCellDisplayText(i, colStockID)
                        .StockBalance = Val(Me.gvSalesDescription.GetRowCellDisplayText(i, colQuantity))
                        .UpdateStock(_DB, _Transaction)


                    End If

                Next



                With CWBCollections
                    .SystemID = Convert.ToInt64(IIf(lblSalesID.Text = String.Empty, 0, lblSalesID.Text))
                    .TransactionTypeID = CWBCommon.CWBEnums.EnumTransactionTypes.SALES
                    .CollectionDelete(_DB, _Transaction)


                    'Add new Records to Collection Table

                    For i As Integer = 0 To Me.gvCollections.RowCount
                        If Not gvCollections.GetRowCellDisplayText(i, gvCollections.Columns(0)) = "" Then
                            .SystemID = Convert.ToInt64(IIf(lblSalesID.Text = String.Empty, 0, lblSalesID.Text))
                            .TransactionTypeID = CWBCommon.CWBEnums.EnumTransactionTypes.SALES

                            Select Case Me.gvCollections.GetRowCellDisplayText(i, GridColumn22)

                                Case "CASH"
                                    .PaymentTypeID = CWBCommon.CWBEnums.EnumPaymentTypes.CASH
                                Case "CHECK"
                                    .PaymentTypeID = CWBCommon.CWBEnums.EnumPaymentTypes.CHECK
                                Case "CR CARD"
                                    .PaymentTypeID = CWBCommon.CWBEnums.EnumPaymentTypes.CCARD

                            End Select

                            .Date = Date.Parse(Me.gvCollections.GetRowCellDisplayText(i, GridColumn25))
                            .Reference = Me.gvCollections.GetRowCellDisplayText(i, GridColumn26)
                            .Amount = FormatNumber(Me.gvCollections.GetRowCellDisplayText(i, GridColumn27), 2, , , )
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
    Private Sub ShowPrintPreview(ByVal SalesId As Int64)
        Try
            Dim _frmPrintPreview As New frmPrint
            Dim _xrSales As New xrSales





            With CWBSales
                .SalesID = SalesId
                .GetSalesByID()

                CWBCustomers.CustomerID = Me.lupCustomer.EditValue
                CWBCustomers.GetCustomerByID()




                _xrSales.xrlCustomers.Text = "[" + CWBCustomers.CustomerNo.ToString + "]  " + CWBCustomers.Salutation + " " + lupCustomer.Text

                _xrSales.xrlAddress1.Text = CWBCustomers.AddressLine1
                _xrSales.xrlAddress2.Text = CWBCustomers.AddressLine2
                _xrSales.xrlAddress3.Text = CWBCustomers.AddressLine3



                _xrSales.xrlSalesNo.Text = CWBSales.SalesNo
                _xrSales.xrlDate.Text = Format(IIf(Not IsDBNull(.SalesDate), .SalesDate, String.Empty), "dd-MMM-yy")
                _xrSales.xrlVehicleNo.Text = IIf(Not IsDBNull(.ReferenceNo), .ReferenceNo, String.Empty)
                _xrSales.xrlModal.Text = leModal.Text

                _xrSales.xrlMeterReading.Text = IIf(Not IsDBNull(.CurrenMeterReading), FormatNumber(.CurrenMeterReading, 0, , , ), "0")
                _xrSales.xrlNextReading.Text = IIf(Not IsDBNull(.NextMeterReading), FormatNumber(.NextMeterReading, 0, , , ), "0")
                _xrSales.xrlKM1.Text = "Km"
                _xrSales.xrlKm2.Text = "Km"


                Dim dtStock As New DataTable
                Dim dtCollection As New DataTable

                dtStock = CWBSales.GetSalesDescriptionByID.Tables(0)

                If dtStock.Rows.Count > 0 Then

                    For i As Integer = 0 To dtStock.Rows.Count - 1
                        Dim tr As New DevExpress.XtraReports.UI.XRTableRow
                        tr.Height = 10

                        Dim cell1, cell2, cell3, cell4, cell5, cell6 As New DevExpress.XtraReports.UI.XRTableCell

                        cell1.Size = New Size(116, 10)
                        cell2.Size = New Size(235, 10)
                        cell3.Size = New Size(107, 10)
                        cell4.Size = New Size(141, 10)
                        'cell5.Size = New Size(125, 25)
                        cell6.Size = New Size(131, 10)


                        If IsDBNull(dtStock.Rows(i).Item("StockCode")) Then
                            cell1.Text = "N/A"
                        Else
                            cell1.Text = CStr((dtStock.Rows(i).Item("StockCode")))
                        End If

                        If IsDBNull(dtStock.Rows(i).Item("Description")) Then
                            cell2.Text = "N/A"
                        Else
                            cell2.Text = dtStock.Rows(i).Item("Description")
                        End If


                        If IsDBNull(dtStock.Rows(i).Item("UnitPrice")) Then
                            cell3.Text = "0.00"
                        Else
                            cell3.Text = CStr(FormatNumber(dtStock.Rows(i).Item("UnitPrice"), 2, , , ))
                        End If



                        If IsDBNull(dtStock.Rows(i).Item("Quantity")) Then
                            cell4.Text = "0.00"
                        Else
                            cell4.Text = CStr(dtStock.Rows(i).Item("Quantity"))
                        End If


                        'If IsDBNull(dtStock.Rows(i).Item("Discount")) Then
                        '    cell5.Text = "0.00"
                        'Else
                        '    cell5.Text = CStr(FormatNumber(dtStock.Rows(i).Item("Discount"), 2, , , ))
                        'End If



                        If IsDBNull(dtStock.Rows(i).Item("Value")) Then
                            cell6.Text = "0.00"
                        Else
                            cell6.Text = CStr(FormatNumber(dtStock.Rows(i).Item("Value"), 2, , , ))
                        End If


                        cell1.TextAlignment = DevExpress.XtraPrinting.TextAlignment.MiddleLeft
                        cell2.TextAlignment = DevExpress.XtraPrinting.TextAlignment.MiddleLeft
                        cell3.TextAlignment = DevExpress.XtraPrinting.TextAlignment.MiddleRight
                        cell4.TextAlignment = DevExpress.XtraPrinting.TextAlignment.MiddleCenter
                        'cell5.TextAlignment = DevExpress.XtraPrinting.TextAlignment.MiddleRight
                        cell6.TextAlignment = DevExpress.XtraPrinting.TextAlignment.MiddleRight



                        tr.Cells.Add(cell1)
                        tr.Cells.Add(cell2)
                        tr.Cells.Add(cell3)
                        tr.Cells.Add(cell4)
                        'tr.Cells.Add(cell5)
                        tr.Cells.Add(cell6)

                        If i = 0 Then
                            _xrSales.xrtStock.Rows.FirstRow.Cells(0).Text = cell1.Text
                            _xrSales.xrtStock.Rows.FirstRow.Cells(1).Text = cell2.Text
                            _xrSales.xrtStock.Rows.FirstRow.Cells(2).Text = cell3.Text
                            _xrSales.xrtStock.Rows.FirstRow.Cells(3).Text = cell4.Text
                            '_xrSales.xrtStock.Rows.FirstRow.Cells(4).Text = cell5.Text
                            _xrSales.xrtStock.Rows.FirstRow.Cells(4).Text = cell6.Text

                        Else
                            _xrSales.xrtStock.Rows.Add(tr)
                        End If

                    Next


                End If



                CWBCollections.SystemID = CWBSales.SalesID
                CWBCollections.TransactionTypeID = CWBCommon.CWBEnums.EnumTransactionTypes.SALES
                dtCollection = CWBCollections.CollectionGetByID().Tables(0)


                'If dtCollection.Rows.Count > 0 Then

                '    For i As Integer = 0 To dtCollection.Rows.Count - 1
                '        Dim tr As New DevExpress.XtraReports.UI.XRTableRow
                '        Dim cell1, cell2, cell3, cell4, cell5, cell6 As New DevExpress.XtraReports.UI.XRTableCell

                '        cell1.Size = New Size(92, 25)
                '        cell2.Size = New Size(108, 25)
                '        cell3.Size = New Size(167, 25)
                '        cell4.Size = New Size(133, 25)



                '        If IsDBNull(dtCollection.Rows(i).Item("Description")) Then
                '            cell1.Text = "N/A"
                '        Else
                '            cell1.Text = CStr((dtCollection.Rows(i).Item("Description")))
                '        End If

                '        If IsDBNull(dtCollection.Rows(i).Item("Date")) Then
                '            cell2.Text = "N/A"
                '        Else
                '            cell2.Text = Format(dtCollection.Rows(i).Item("Date"), "dd-MMM-yy")
                '        End If



                '        If IsDBNull(dtCollection.Rows(i).Item("Reference")) Then
                '            cell3.Text = "0.00"
                '        Else
                '            cell3.Text = CStr(dtCollection.Rows(i).Item("Reference"))
                '        End If



                '        If IsDBNull(dtCollection.Rows(i).Item("Amount")) Then
                '            cell4.Text = "0.00"
                '        Else
                '            cell4.Text = CStr(FormatNumber(dtCollection.Rows(i).Item("Amount"), 2, , , ))
                '        End If





                '        cell1.TextAlignment = DevExpress.XtraPrinting.TextAlignment.MiddleLeft
                '        cell2.TextAlignment = DevExpress.XtraPrinting.TextAlignment.MiddleLeft
                '        cell3.TextAlignment = DevExpress.XtraPrinting.TextAlignment.MiddleLeft
                '        cell4.TextAlignment = DevExpress.XtraPrinting.TextAlignment.MiddleRight



                '        tr.Cells.Add(cell1)
                '        tr.Cells.Add(cell2)
                '        tr.Cells.Add(cell3)
                '        tr.Cells.Add(cell4)


                '        'If i = 0 Then
                '        '    _xrSales.xrtCollections.Rows.FirstRow.Cells(0).Text = cell1.Text
                '        '    _xrSales.xrtCollections.Rows.FirstRow.Cells(1).Text = cell2.Text
                '        '    _xrSales.xrtCollections.Rows.FirstRow.Cells(2).Text = cell3.Text
                '        '    _xrSales.xrtCollections.Rows.FirstRow.Cells(3).Text = cell4.Text


                '        'Else
                '        '    _xrSales.xrtCollections.Rows.Add(tr)
                '        'End If



                '    Next
                'End If

                '_xrSales.xrlTotal.Text = FormatNumber(IIf(Not IsDBNull(.Total), .Total, 0), 2, , , )
                '_xrSales.xrlTax.Text = FormatNumber(IIf(Not IsDBNull(.TaxAmount), .TaxAmount, 0), 2, , , )
                '_xrSales.xrlDiscount.Text = FormatNumber(IIf(Not IsDBNull(.Discount), .Discount, 0), 2, , , )

                Dim a, b, c, d As Decimal
                a = .Total
                'b = .TaxAmount
                c = .Discount

                GridColumn27.SummaryItem.DisplayFormat = ""
                b = gvCollections.Columns("Amount").SummaryText
                GridColumn27.SummaryItem.DisplayFormat = "{0:n2}"


                d = a - (b + c)

                _xrSales.xrlTotal.Text = FormatNumber(a, 2, , , )
                _xrSales.xrlPaid.Text = FormatNumber(b, 2, , , )

                _xrSales.xrlDue.Text = FormatNumber(d, 2, , , )


                If CWBSales.Ch1 = True Then

                    _xrSales.xrlCh1.Text = "CH"

                Else
                    _xrSales.xrlCh1.Text = String.Empty
                End If



                If CWBSales.Ch2 = True Then

                    _xrSales.xrlCh2.Text = "CH"

                Else
                    _xrSales.xrlCh2.Text = String.Empty
                End If



                If CWBSales.Ch3 = True Then

                    _xrSales.xrlCh3.Text = "CH"

                Else
                    _xrSales.xrlCh3.Text = String.Empty
                End If




                If CWBSales.Ch4 = True Then

                    _xrSales.xrlCh4.Text = "CH"

                Else
                    _xrSales.xrlCh4.Text = String.Empty
                End If

                If leMessages.Text = "NONE" Then
                    _xrSales.xrlMessage.Text = String.Empty
                Else
                    _xrSales.xrlMessage.Text = leMessages.Text
                End If
                
                If CWBSales.TaxPercent = 0 Then
                    _xrSales.xrlDiscount.Text = FormatNumber(CWBSales.Discount, 2, , , )

                Else
                    _xrSales.xrlDiscount.Text = "(" + FormatNumber(CWBSales.TaxPercent, 2, , , ) + "%)  " + FormatNumber(CWBSales.Discount, 2, , , )

                End If


            End With




            _frmPrintPreview.PrintControl1.PrintingSystem = _xrSales.PrintingSystem
            _xrSales.CreateDocument()
            _frmPrintPreview.MdiParent = frmMain
            _frmPrintPreview.Show()
            _frmPrintPreview.BringToFront()

        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Model Lookup Events"
    Private Sub leModal_ButtonClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraEditors.Controls.ButtonPressedEventArgs) Handles leModal.ButtonClick
        Select Case e.Button.Index
            Case 1
                If Not Me.leModal.Text = String.Empty Then
                    Me.SaveModels()
                    Me.PopulateModelLookup()
                End If

            Case 2
                If Not Me.leModal.Text = String.Empty Then
                    Dim frm As New frmDeleteYesNo
                    frm.lblTitle.Text = CWB_DELETE_CONFIRMATION_TITLELABEL
                    frm.lblDescription.ForeColor = Color.Red
                    frm.peImage.Image = CWB.My.Resources.Resources.ImgDelete
                    frm.lblDescription.Text = CWB_DELETE_CONFIRMATION_DESCRIPTIONLABEL
                    frm.Text = CWB_DELETE_CONFIRMATION_TITLE


                    If frm.ShowDialog = Windows.Forms.DialogResult.Yes Then
                        Me.DeleteModels()
                        Me.PopulateModelLookup()
                    End If
                End If



        End Select
    End Sub
#End Region

#Region "Save Model"
    Private Sub SaveModels()
        Try
            With CWBModels
                .Description = Me.leModal.Text
                .InsertModels()

            End With
        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Save Message"
    Private Sub SaveMessage()
        Try
            With CWBMessages
                .MessageText = Me.leMessages.Text
                .InsertMessage()

            End With
        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Populate Model Look up"
    Public Sub PopulateModelLookup()

        Try
            With leModal
                .Properties.DataSource = CWBModels.GetAllModels.Tables(0)
                .Properties.DisplayMember = "Description"
                .Properties.ValueMember = "ModelID"

            End With


        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Populate Messages"
    Public Sub PopulateMessageLookup()

        Try
            With leMessages
                .Properties.DataSource = CWBMessages.GetAllMessage.Tables(0)
                .Properties.DisplayMember = "Message"
                .Properties.ValueMember = "MessageID"

            End With


        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Delete Models"
    Private Sub DeleteModels()
        Try
            With CWBModels
                .Description = Me.leModal.Text
                .DeleteModels()
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

#Region "Delete Message"
    Private Sub DeleteMessage()
        Try
            With CWBMessages
                .MessageText = Me.leMessages.Text
                .DeleteMessage()
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

#Region "Calculate Balance"
    Private Sub CalculateBalance()
        Try

            Dim total, taxamount, discount, balance As Decimal



            GridColumn40.SummaryItem.DisplayFormat = ""
            total = Val(GridColumn40.SummaryText)
            GridColumn40.SummaryItem.DisplayFormat = "Total : {0:n2}"


            taxamount = Me.seTaxAmount.EditValue
            discount = Me.seDiscount.EditValue
            balance = total + taxamount - (discount)

            seGrandTotal.EditValue = balance

        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Text Edit Events"
    Private Sub teReferenceNo_EditValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles teReferenceNo.EditValueChanged

        CWBSales.ReferenceNo = Me.teReferenceNo.Text

        If CWBSales.ReferenceNo = String.Empty Then
            lcPreviousBills.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never
        Else
            If CWBSales.GetSalesByReferenceNo.Tables(0).Rows.Count > 0 Then
                lcPreviousBills.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always
            Else
                lcPreviousBills.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never
            End If

        End If

       






    End Sub
#End Region

#Region "Button Events"
    Private Sub sbPreviousBills_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sbPreviousBills.Click
        Try
            CWBSales.ReferenceNo = Me.teReferenceNo.Text.Trim
            frmPreviousBills.lblDescription.Text = "Previous Bills for " + Me.teReferenceNo.Text.Trim


            Dim dset As New DataSet

            dset = CWBSales.GetSalesByReferenceNo



            Dim keyColumn As DataColumn = dset.Tables(0).Columns("SalesID")
            Dim foreignKeyColumn As DataColumn = dset.Tables(1).Columns("SalesID")
            dset.Relations.Add("Bill Description", keyColumn, foreignKeyColumn)
            frmPreviousBills.gcPreviousBills.LevelTree.Nodes.Add("Bill Description", frmPreviousBills.gvDescription)
            frmPreviousBills.gcPreviousBills.ForceInitialize()
            frmPreviousBills.gcPreviousBills.DataSource = dset.Tables(0)

            frmPreviousBills.ShowDialog()




        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Button Events"
    Private Sub sbProcess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sbProcess.Click
        Try
            If dxvpHistoryData.Validate Then
                PopulateSalesHistory()
            End If


        Catch ex As Exception

            MessageError(ex.ToString)
        End Try

    End Sub
#End Region

#Region "Lookup  Event"
    Private Sub leService1_ButtonClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraEditors.Controls.ButtonPressedEventArgs) Handles leService1.ButtonClick

        Try
            Select Case e.Button.Index

                Case 1
                    leService1.EditValue = Nothing
            End Select
        Catch ex As Exception
            MessageError(ex.ToString)
        End Try

    End Sub
    Private Sub leService2_ButtonClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraEditors.Controls.ButtonPressedEventArgs) Handles leService2.ButtonClick
        Try
            Select Case e.Button.Index

                Case 1
                    leService2.EditValue = Nothing
            End Select
        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
    Private Sub leMessages_ButtonClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraEditors.Controls.ButtonPressedEventArgs) Handles leMessages.ButtonClick
        Select Case e.Button.Index
            Case 1
                If Not Me.leMessages.Text = String.Empty Then
                    Me.SaveMessage()
                    Me.PopulateMessageLookup()
                End If

            Case 2
                If Not Me.leMessages.Text = String.Empty Then
                    Dim frm As New frmDeleteYesNo
                    frm.lblTitle.Text = CWB_DELETE_CONFIRMATION_TITLELABEL
                    frm.lblDescription.ForeColor = Color.Red
                    frm.peImage.Image = CWB.My.Resources.Resources.ImgDelete
                    frm.lblDescription.Text = CWB_DELETE_CONFIRMATION_DESCRIPTIONLABEL
                    frm.Text = CWB_DELETE_CONFIRMATION_TITLE


                    If frm.ShowDialog = Windows.Forms.DialogResult.Yes Then
                        Me.DeleteMessage()
                        Me.PopulateMessageLookup()
                    End If
                End If



        End Select
    End Sub
#End Region

#Region "Sping Edit Evets"

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

#Region "Calculate Discount Percent"
    Private Sub CalculateDiscountPercent()

        Try

            Dim total, discount As Decimal



            GridColumn40.SummaryItem.DisplayFormat = ""
            total = Val(GridColumn40.SummaryText)
            GridColumn40.SummaryItem.DisplayFormat = "Total : {0:n2}"




            discount = total * (sePercentage.EditValue / 100)


            seDiscount.EditValue = discount



        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub

    Private Sub bbAudit_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbAudit.ItemClick

        If (Me.lblSystemNo.Text <> String.Empty) Then
            frmAudit.lblSalesId.Text = Me.lblSystemNo.Text
            frmAudit.ShowDialog()
        End If


    End Sub
#End Region


End Class