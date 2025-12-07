Imports CWBCommon.CWBStock
Imports CWBCommon.CWBSuppliers
Imports CWBCommon.CWBConstants


Public Class frmStock

#Region "Variables"
    Private _CWBStock As CWBCommon.CWBStock
    Private _CWBSuppliers As CWBCommon.CWBSuppliers
#End Region

#Region "Constructor"

    Public ReadOnly Property CWBStock() As CWBCommon.CWBStock
        Get

            If _CWBStock Is Nothing Then
                _CWBStock = New CWBCommon.CWBStock()
            End If

            Return _CWBStock
        End Get
    End Property

    Public ReadOnly Property CWBSuppliers() As CWBCommon.CWBSuppliers
        Get

            If _CWBSuppliers Is Nothing Then
                _CWBSuppliers = New CWBCommon.CWBSuppliers()
            End If

            Return _CWBSuppliers
        End Get
    End Property
#End Region

#Region "Form Events"

    Private Sub frmStockItems_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Me.PopulateSupplierLookup()
        Me.PopulateCategoryLookup()
    End Sub

    Private Sub frmStockItems_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.Chr(27) Then
            Me.Close()
        End If
    End Sub

    Private Sub frmStockItems_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetStartUpBarButton(bbSave, bbNew, bbDelete, bbClose, bbDisplaySelected, bbRefresh, bbPrint)
        Me.HideToolButtonsOnLoad()
    End Sub

#End Region

#Region "Bar button Event"
    Private Sub bbSave_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbSave.ItemClick
        If dxvpStockItems.Validate Then
            If lblID.Text = String.Empty Then

                CWBStock.StockCode = Me.teStockCode.Text


                If CWBStock.IsStockExists() = True Then
                    Dim frm As New frmAlreadyExists

                    frm.Text = CWB_SUPPLIER_CONFIRMATION_TITLE
                    frm.lblTitle.Text = CWB_SUPPLIER_CONFIRMATION_TITLELABEL
                    frm.lblDescription.Text = CWB_SUPPLIER_CONFIRMATION_DESCRIPTIONLABEL
                    frm.ShowDialog()
                    Me.teStockCode.Focus()
                Else
                    Me.SaveRecords()
                End If


            Else

                Dim frm As New frmUpdateYesNo
                frm.peImage.Image = CWB.My.Resources.Resources.ImgUpdate
                frm.Text = CWB_UPDATE_CONFIRMATION_TITLE
                frm.lblTitle.Text = CWB_UPDATE_CONFIRMATION_TITLELABEL
                frm.lblDescription.Text = CWB_UPDATE_CONFIRMATION_DESCRIPTIONLABEL

                If UserType = "OWNER" Then
                    If frm.ShowDialog = Windows.Forms.DialogResult.Yes Then

                        CWBStock.StockCode = Me.teStockCode.Text
                        CWBStock.GetByStockCode()

                        If CWBStock.StockID = Me.lblID.Text Then
                            Me.SaveRecords()
                        Else

                            Dim frmx As New frmAlreadyExists

                            frmx.Text = CWB_STOCK_CONFIRMATION_TITLE
                            frmx.lblTitle.Text = CWB_STOCK_CONFIRMATION_TITLELABEL
                            frmx.lblDescription.Text = CWB_STOCK_CONFIRMATION_DESCRIPTIONLABEL
                            frmx.ShowDialog()
                            Me.teStockCode.Focus()

                        End If

                    End If
                Else
                    MessageOK("Can not Update", "You don't have permission", "Click OK to continue")
                End If


              


                

            End If

            End If

    End Sub

    Private Sub bbPrint_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbPrint.ItemClick
        PrintPreview(gcStockItems, "Stock Report")

    End Sub

    Private Sub bbNew_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbNew.ItemClick
        Me.ClearFormData()
    End Sub

    Private Sub bbDelete_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbDelete.ItemClick
        If Not teStockCode.EditValue = String.Empty Then
            Dim frm As New frmDeleteYesNo
            frm.lblTitle.Text = CWB_DELETE_CONFIRMATION_TITLELABEL
            frm.lblDescription.ForeColor = Color.Red
            frm.peImage.Image = CWB.My.Resources.Resources.ImgDelete
            frm.lblDescription.Text = CWB_DELETE_CONFIRMATION_DESCRIPTIONLABEL
            frm.Text = CWB_DELETE_CONFIRMATION_TITLE

            If frm.ShowDialog = Windows.Forms.DialogResult.Yes Then

                Me.DeleteStock(teStockCode.EditValue)
                Me.ClearFormData()
            End If
        End If

        teStockCode.Focus()
    End Sub

    Private Sub bbClose_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbClose.ItemClick
        Me.Close()
    End Sub

    Private Sub bbDisplaySelected_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbDisplaySelected.ItemClick
        Me.DisplayStock()
    End Sub

    Private Sub bbRefresh_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbRefresh.ItemClick
        With gvStockItems
            .ClearColumnsFilter()
            .ClearGrouping()
            .ClearSelection()
            .ClearSorting()
        End With

    End Sub
#End Region

#Region "Save Stock Data"
    Private Sub SaveRecords()
        Try
            With CWBStock
                .StockID = Convert.ToInt64(IIf(lblID.Text = String.Empty, 0, lblID.Text))
                .StockCode = teStockCode.Text.Trim
                .Description = teDescription.Text.Trim
                .StockCategoryID = LupStockCategory.EditValue
                .StockType = ceStockType.EditValue
                .PurchasingPrice = sePurchasePrice.EditValue
                .SalesPrice = seSalesPrice.EditValue
                .SupplierID = lupSupplier.EditValue
                .ReorderLevel = seReorderLevel.EditValue
                .OpeningStockBalance = seOpeningStockBalance.EditValue
                .OpeningStockValue = seOpeningStockValue.EditValue
                .CurrentStockBalance = seCurrentStockBalance.EditValue
                .CurrentStockValue = seCurrentStockValue.EditValue
                .CreatedBy = UserID
                .UpdatedBy = UserID
                .InsertStock()
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

#Region "Save Stock Category"
    Private Sub SaveStockCategory()
        Try
            With CWBStock
                .StockCategoryName = Me.LupStockCategory.Text
                .InsertStockCategory()

            End With
        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Delete Stock Category"
    Private Sub DeleteStockCategory()
        Try
            With CWBStock
                .StockCategoryName = Me.LupStockCategory.Text
                .DeleteStockCategories()
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

#Region "Clears Form Data"

    Private Sub ClearFormData()

        Try
            lblID.Text = String.Empty
            teStockCode.Text = String.Empty
            teDescription.Text = String.Empty
            sePurchasePrice.Text = 0
            seSalesPrice.Text = 0
            seReorderLevel.Text = 0
            seOpeningStockBalance.Text = 0
            seOpeningStockValue.Text = 0
            seCurrentStockBalance.Text = 0
            seCurrentStockValue.Text = 0
            lblID.Text = String.Empty
            LupStockCategory.EditValue = DBNull.Value
            lupSupplier.EditValue = DBNull.Value

            dxvpStockItems.RemoveControlError(teStockCode)
            dxvpStockItems.RemoveControlError(teDescription)
            dxvpStockItems.RemoveControlError(LupStockCategory)
            dxvpStockItems.RemoveControlError(lupSupplier)
            teStockCode.Focus()

        Catch ex As Exception
            MessageError(ex.ToString)
        End Try


    End Sub

#End Region

#Region "Display Stock"
    Public Sub DisplayStock()

        Try
            If gvStockItems.RowCount > 0 Then
                xTab1.SelectedTabPageIndex = 0
                With CWBStock
                    .StockID = gvStockItems.GetFocusedRowCellValue(GridColumn1)
                    .GetByStockID()
                    Me.lblID.Text = .StockID
                    teStockCode.EditValue = .StockCode
                    teDescription.EditValue = .Description
                    ceStockType.EditValue = .StockType
                    LupStockCategory.EditValue = .StockCategoryID
                    sePurchasePrice.EditValue = .PurchasingPrice
                    seSalesPrice.EditValue = .SalesPrice
                    lupSupplier.EditValue = .SupplierID
                    seReorderLevel.EditValue = .ReorderLevel
                    seOpeningStockBalance.EditValue = .OpeningStockBalance
                    seOpeningStockValue.EditValue = .OpeningStockValue
                    seCurrentStockBalance.EditValue = .CurrentStockBalance
                    seCurrentStockValue.EditValue = .CurrentStockValue
                End With
            End If

        Catch ex As Exception
            Dim frm As New frmOk
            frm.Text = CWB_DELETE_ERROR_CONFIRMATION_TITLE
            frm.lblTitle.Text = CWB_DELETE_ERROR_CONFIRMATION_TITLELABEL
            frm.lblDescription.Text = CWB_DELETE_ERROR_CONFIRMATION_DESCRIPTIONLABEL
            frm.ShowDialog()
        End Try




    End Sub
#End Region

#Region "Delete Stock"
    Private Sub DeleteStock(ByVal strStockCode As String)
        Try
            With CWBStock
                .StockCode = strStockCode
                .DeleteStock()
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

#Region "Hide Tool Buttons On Load"
    Public Sub HideToolButtonsOnLoad()
        Me.bbShowAll.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
        Me.bbDisplaySelected.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
        Me.bbRefresh.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
        Me.bbPrint.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
    End Sub
#End Region

#Region "Show Tool Buttons On History Tab change"
    Public Sub ShowToolButtonsOnHistoryTabChange()
        Me.bbShowAll.Visibility = DevExpress.XtraBars.BarItemVisibility.Always
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
        Me.bbShowAll.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
        Me.bbDisplaySelected.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
        Me.bbRefresh.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
        Me.bbPrint.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
        Me.bbSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Always
        Me.bbNew.Visibility = DevExpress.XtraBars.BarItemVisibility.Always
        Me.bbDelete.Visibility = DevExpress.XtraBars.BarItemVisibility.Always
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
    Private Sub gvStockItems_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gvStockItems.DoubleClick
        Me.DisplayStock()
    End Sub
#End Region

#Region "Populate Grid"
    Public Sub PopulateGrid()
        Try
            gcStockItems.DataSource = CWBStock.GetAllStockItems().Tables(0)
        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Populate Suppliers Lookup"
    Public Sub PopulateSupplierLookup()

        Try
            With lupSupplier
                .Properties.DataSource = CWBSuppliers.GetAllSuppliers.Tables(0)
                .Properties.DisplayMember = "SupplierName"
                .Properties.ValueMember = "SupplierID"

            End With


        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Populate Stock Category Lookup"
    Public Sub PopulateCategoryLookup()

        Try
            With LupStockCategory
                .Properties.DataSource = CWBStock.GetAllStockCategories.Tables(0)
                .Properties.DisplayMember = "Name"
                .Properties.ValueMember = "StockCategoryID"

            End With


        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Lookup Events"
    Private Sub LupStockCategory_ButtonClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraEditors.Controls.ButtonPressedEventArgs) Handles LupStockCategory.ButtonClick
        Select Case e.Button.Index
            Case 1
                If Not Me.LupStockCategory.Text = String.Empty Then
                    Me.SaveStockCategory()
                    Me.PopulateCategoryLookup()
                End If

            Case 2
                If Not Me.LupStockCategory.Text = String.Empty Then
                    Dim frm As New frmDeleteYesNo
                    frm.lblTitle.Text = CWB_DELETE_CONFIRMATION_TITLELABEL
                    frm.lblDescription.ForeColor = Color.Red
                    frm.peImage.Image = CWB.My.Resources.Resources.ImgDelete
                    frm.lblDescription.Text = CWB_DELETE_CONFIRMATION_DESCRIPTIONLABEL
                    frm.Text = CWB_DELETE_CONFIRMATION_TITLE

                    If frm.ShowDialog = Windows.Forms.DialogResult.Yes Then
                        Me.DeleteStockCategory()
                        Me.PopulateCategoryLookup()
                    End If
                End If



        End Select
    End Sub
#End Region

#Region "Spin Events"
    Private Sub seOpeningStockValue_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles seOpeningStockValue.Enter
        Me.CalculateOpeningValue()
    End Sub

    Private Sub sePurchasePrice_EditValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sePurchasePrice.EditValueChanged
        Me.CalculateOpeningValue()
    End Sub

    Private Sub seOpeningStockBalance_EditValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles seOpeningStockBalance.EditValueChanged
        Me.CalculateOpeningValue()
    End Sub
#End Region

#Region "Calculate Opening Value"
    Private Sub CalculateOpeningValue()
        Try

            Dim purchaseprice, openingbalance, openingvalue As Decimal

            purchaseprice = Val(Me.sePurchasePrice.EditValue)
            openingbalance = Val(seOpeningStockBalance.EditValue)
            openingvalue = purchaseprice * openingbalance
            Me.seOpeningStockValue.EditValue = openingvalue


        Catch ex As Exception

            MessageError(ex.ToString)
        End Try
    End Sub
#End Region


End Class