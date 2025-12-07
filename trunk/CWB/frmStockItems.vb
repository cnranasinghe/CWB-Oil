Imports CWBCommon.CWBStockItems
Imports CWBCommon.CWBSuppliers


Public Class frmStockItems

#Region "Variables"
    Private _CWBStock As CWBCommon.CWBStockItems
    Private _CWBSuppliers As CWBCommon.CWBSuppliers
#End Region

#Region "Constructor"

    Public ReadOnly Property DAStock() As CWBCommon.CWBStockItems
        Get

            If _CWBStock Is Nothing Then
                _CWBStock = New CWBCommon.CWBStockItems()
            End If

            Return _CWBStock
        End Get
    End Property

    Public ReadOnly Property DASuppliers() As CWBCommon.CWBSuppliers
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
                Me.SaveRecords()
            Else
                Dim frm As New frmUpdateYesNo
                frm.Text = "Record Already exists"
                frm.lblTitle.Text = "Record Already Exists"
                frm.lblDescription.Text = "Do you want to replace the selected record?"

                If frm.ShowDialog = Windows.Forms.DialogResult.Yes Then
                    Me.SaveRecords()
                End If
            End If

        End If

    End Sub

    Private Sub bbNew_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbNew.ItemClick
        Me.ClearFormData()
    End Sub

    Private Sub bbDelete_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbDelete.ItemClick
        If Not teStockCode.EditValue = String.Empty Then
            Dim f As New frmUpdateYesNo
            f.lblTitle.Text = "It is not recormended to delete the selected record?"
            f.lblDescription.ForeColor = Color.Red
            f.lblDescription.Text = "Click Yes to Delete..."


            f.Text = "Delete The Current Record ?"

            If f.ShowDialog = Windows.Forms.DialogResult.Yes Then

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
        With DAStock
            .StockID = Convert.ToInt64(IIf(lblID.Text = String.Empty, 0, lblID.Text))
            .Stock_Code = teStockCode.Text.Trim
            .Description = teDescription.Text.Trim
            .Stock_Category = LupStockCategory.EditValue
            .Purchasing_Price = sePurchasePrice.EditValue
            .Sales_Price = seSalesPrice.EditValue
            .Supplier = lupSupplier.EditValue
            .Reorder_Level = seReorderLevel.EditValue
            .Stock_Balance = seStockBalance.EditValue
            .Stock_Value = seStockValue.EditValue
            .InsertStock()
        End With
        Me.ClearFormData()
    End Sub
#End Region

#Region "Save Stock Category"
    Private Sub SaveStockCategory()
        Try
            With DAStock
                .StockCategoryName = Me.LupStockCategory.Text
                .InsertStockCategory()

            End With
        Catch ex As Exception

        End Try
    End Sub
#End Region

#Region "Delete Stock Category"
    Private Sub DeleteStockCategory()
        Try
            With DAStock
                .StockCategoryName = Me.LupStockCategory.Text
                .DeleteStockCategories()
            End With
        Catch ex As Exception

        End Try
    End Sub
#End Region

#Region "Clears Form Data"

    Private Sub ClearFormData()

        teStockCode.Text = String.Empty
        teDescription.Text = String.Empty
        sePurchasePrice.Text = 0
        seSalesPrice.Text = 0
        seReorderLevel.Text = 0
        seStockBalance.Text = 0
        seStockValue.Text = 0
        lblID.Text = String.Empty
        LupStockCategory.EditValue = DBNull.Value
        lupSupplier.EditValue = DBNull.Value
        teStockCode.Focus()

    End Sub

#End Region

#Region "Display Stock"
    Public Sub DisplayStock()

        If gvStockItems.RowCount > 0 Then
            xTab1.SelectedTabPageIndex = 0
            With DAStock
                .StockID = gvStockItems.GetFocusedRowCellValue(GridColumn1)
                .GetByStockID()
                Me.lblID.Text = .StockID
                teStockCode.EditValue = .Stock_Code
                teDescription.EditValue = .Description
                LupStockCategory.EditValue = .Stock_Category
                sePurchasePrice.EditValue = .Purchasing_Price
                seSalesPrice.EditValue = .Sales_Price
                lupSupplier.EditValue = .Supplier
                seReorderLevel.EditValue = .Reorder_Level
                seStockBalance.EditValue = .Stock_Balance
                seStockValue.EditValue = .Stock_Value
            End With
        End If



    End Sub
#End Region

#Region "Delete Stock"
    Private Sub DeleteStock(ByVal strStockCode As String)
        With DAStock
            .Stock_Code = strStockCode
            .DeleteStock()
        End With


    End Sub
#End Region

#Region "Hide Tool Buttons On Load"
    Public Sub HideToolButtonsOnLoad()
        Me.bbShowAll.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
        Me.bbDisplaySelected.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
        Me.bbRefresh.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
    End Sub
#End Region

#Region "Show Tool Buttons On History Tab change"
    Public Sub ShowToolButtonsOnHistoryTabChange()
        Me.bbShowAll.Visibility = DevExpress.XtraBars.BarItemVisibility.Always
        Me.bbDisplaySelected.Visibility = DevExpress.XtraBars.BarItemVisibility.Always
        Me.bbRefresh.Visibility = DevExpress.XtraBars.BarItemVisibility.Always
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
            gcStockItems.DataSource = DAStock.GetAllStockItems().Tables(0)
        Catch ex As Exception

        End Try
    End Sub
#End Region

#Region "Populate Suppliers Lookup"
    Public Sub PopulateSupplierLookup()

        Try
            With lupSupplier
                .Properties.DataSource = DASuppliers.GetAllSuppliers.Tables(0)
                .Properties.DisplayMember = "Name"
                .Properties.ValueMember = "SupplierID"

            End With


        Catch ex As Exception

        End Try
    End Sub
#End Region

#Region "Populate Stock Category Lookup"
    Public Sub PopulateCategoryLookup()

        Try
            With LupStockCategory
                .Properties.DataSource = DAStock.GetAllStockCategories.Tables(0)
                .Properties.DisplayMember = "Name"
                .Properties.ValueMember = "StockCategoryID"

            End With


        Catch ex As Exception

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
                    Dim f As New frmUpdateYesNo
                    f.lblTitle.Text = "It is not recormended to delete the selected record?"
                    f.lblDescription.ForeColor = Color.Red
                    f.lblDescription.Text = "Click Yes to Delete..."
                    f.Text = "Delete The Current Record ?"

                    If f.ShowDialog = Windows.Forms.DialogResult.Yes Then
                        Me.DeleteStockCategory()
                        Me.PopulateCategoryLookup()
                    End If
                End If



        End Select
    End Sub
#End Region

End Class