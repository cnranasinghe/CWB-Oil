<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDailySummary
    Inherits DevExpress.XtraEditors.XtraForm

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim CompareAgainstControlValidationRule1 As DevExpress.XtraEditors.DXErrorProvider.CompareAgainstControlValidationRule = New DevExpress.XtraEditors.DXErrorProvider.CompareAgainstControlValidationRule
        Me.deFromDate = New DevExpress.XtraEditors.DateEdit
        Me.LayoutControl1 = New DevExpress.XtraLayout.LayoutControl
        Me.xtab1 = New DevExpress.XtraTab.XtraTabControl
        Me.XtraTabPage1 = New DevExpress.XtraTab.XtraTabPage
        Me.LayoutControl2 = New DevExpress.XtraLayout.LayoutControl
        Me.gcPurchases = New DevExpress.XtraGrid.GridControl
        Me.gvPurchases = New DevExpress.XtraGrid.Views.Grid.GridView
        Me.GridColumn1 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn2 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn3 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn4 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn5 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn6 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn7 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn8 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn9 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn10 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.LayoutControlGroup2 = New DevExpress.XtraLayout.LayoutControlGroup
        Me.LayoutControlItem1 = New DevExpress.XtraLayout.LayoutControlItem
        Me.XtraTabPage2 = New DevExpress.XtraTab.XtraTabPage
        Me.LayoutControl3 = New DevExpress.XtraLayout.LayoutControl
        Me.gcSales = New DevExpress.XtraGrid.GridControl
        Me.gvSales = New DevExpress.XtraGrid.Views.Grid.GridView
        Me.GridColumn11 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn12 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn13 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn14 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn15 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn16 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn17 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn18 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn19 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn20 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn37 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn38 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.LayoutControlGroup3 = New DevExpress.XtraLayout.LayoutControlGroup
        Me.LayoutControlItem6 = New DevExpress.XtraLayout.LayoutControlItem
        Me.XtraTabPage3 = New DevExpress.XtraTab.XtraTabPage
        Me.LayoutControl4 = New DevExpress.XtraLayout.LayoutControl
        Me.gcExpenses = New DevExpress.XtraGrid.GridControl
        Me.gvExpenses = New DevExpress.XtraGrid.Views.Grid.GridView
        Me.GridColumn21 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn22 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn23 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn24 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.LayoutControlGroup5 = New DevExpress.XtraLayout.LayoutControlGroup
        Me.LayoutControlItem7 = New DevExpress.XtraLayout.LayoutControlItem
        Me.XtraTabPage4 = New DevExpress.XtraTab.XtraTabPage
        Me.LayoutControl5 = New DevExpress.XtraLayout.LayoutControl
        Me.gcStockItems = New DevExpress.XtraGrid.GridControl
        Me.gvStockItems = New DevExpress.XtraGrid.Views.Grid.GridView
        Me.GridColumn25 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn26 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn29 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn30 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.LayoutControlGroup6 = New DevExpress.XtraLayout.LayoutControlGroup
        Me.LayoutControlItem8 = New DevExpress.XtraLayout.LayoutControlItem
        Me.XtraTabPage5 = New DevExpress.XtraTab.XtraTabPage
        Me.LayoutControl6 = New DevExpress.XtraLayout.LayoutControl
        Me.gcServices = New DevExpress.XtraGrid.GridControl
        Me.gvServices = New DevExpress.XtraGrid.Views.Grid.GridView
        Me.GridColumn31 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn32 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn35 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn36 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.LayoutControlGroup7 = New DevExpress.XtraLayout.LayoutControlGroup
        Me.LayoutControlItem9 = New DevExpress.XtraLayout.LayoutControlItem
        Me.sbPrint = New DevExpress.XtraEditors.SimpleButton
        Me.sbProcess = New DevExpress.XtraEditors.SimpleButton
        Me.deToDate = New DevExpress.XtraEditors.DateEdit
        Me.LayoutControlGroup1 = New DevExpress.XtraLayout.LayoutControlGroup
        Me.EmptySpaceItem1 = New DevExpress.XtraLayout.EmptySpaceItem
        Me.EmptySpaceItem2 = New DevExpress.XtraLayout.EmptySpaceItem
        Me.LayoutControlGroup4 = New DevExpress.XtraLayout.LayoutControlGroup
        Me.From = New DevExpress.XtraLayout.LayoutControlItem
        Me.LayoutControlItem4 = New DevExpress.XtraLayout.LayoutControlItem
        Me.LayoutControlItem5 = New DevExpress.XtraLayout.LayoutControlItem
        Me.LayoutControlItem2 = New DevExpress.XtraLayout.LayoutControlItem
        Me.EmptySpaceItem4 = New DevExpress.XtraLayout.EmptySpaceItem
        Me.LayoutControlItem3 = New DevExpress.XtraLayout.LayoutControlItem
        Me.dxvpSalesSummary = New DevExpress.XtraEditors.DXErrorProvider.DXValidationProvider(Me.components)
        CType(Me.deFromDate.Properties.VistaTimeProperties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.deFromDate.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.LayoutControl1.SuspendLayout()
        CType(Me.xtab1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.xtab1.SuspendLayout()
        Me.XtraTabPage1.SuspendLayout()
        CType(Me.LayoutControl2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.LayoutControl2.SuspendLayout()
        CType(Me.gcPurchases, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gvPurchases, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlGroup2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.XtraTabPage2.SuspendLayout()
        CType(Me.LayoutControl3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.LayoutControl3.SuspendLayout()
        CType(Me.gcSales, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gvSales, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlGroup3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem6, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.XtraTabPage3.SuspendLayout()
        CType(Me.LayoutControl4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.LayoutControl4.SuspendLayout()
        CType(Me.gcExpenses, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gvExpenses, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlGroup5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem7, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.XtraTabPage4.SuspendLayout()
        CType(Me.LayoutControl5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.LayoutControl5.SuspendLayout()
        CType(Me.gcStockItems, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gvStockItems, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlGroup6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem8, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.XtraTabPage5.SuspendLayout()
        CType(Me.LayoutControl6, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.LayoutControl6.SuspendLayout()
        CType(Me.gcServices, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gvServices, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlGroup7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem9, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.deToDate.Properties.VistaTimeProperties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.deToDate.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlGroup1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.EmptySpaceItem1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.EmptySpaceItem2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlGroup4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.From, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.EmptySpaceItem4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dxvpSalesSummary, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'deFromDate
        '
        Me.deFromDate.EditValue = Nothing
        Me.deFromDate.EnterMoveNextControl = True
        Me.deFromDate.Location = New System.Drawing.Point(254, 28)
        Me.deFromDate.Name = "deFromDate"
        Me.deFromDate.Properties.AllowNullInput = DevExpress.Utils.DefaultBoolean.[False]
        Me.deFromDate.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.deFromDate.Properties.DisplayFormat.FormatString = "dd-MMM-yy"
        Me.deFromDate.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime
        Me.deFromDate.Properties.EditFormat.FormatString = "dd-MMM-yy"
        Me.deFromDate.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.DateTime
        Me.deFromDate.Properties.Mask.EditMask = "dd-MMM-yy"
        Me.deFromDate.Properties.VistaTimeProperties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton})
        Me.deFromDate.Size = New System.Drawing.Size(160, 20)
        Me.deFromDate.StyleController = Me.LayoutControl1
        Me.deFromDate.TabIndex = 0
        '
        'LayoutControl1
        '
        Me.LayoutControl1.Controls.Add(Me.xtab1)
        Me.LayoutControl1.Controls.Add(Me.sbPrint)
        Me.LayoutControl1.Controls.Add(Me.sbProcess)
        Me.LayoutControl1.Controls.Add(Me.deToDate)
        Me.LayoutControl1.Controls.Add(Me.deFromDate)
        Me.LayoutControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LayoutControl1.Location = New System.Drawing.Point(0, 0)
        Me.LayoutControl1.Name = "LayoutControl1"
        Me.LayoutControl1.Root = Me.LayoutControlGroup1
        Me.LayoutControl1.Size = New System.Drawing.Size(961, 494)
        Me.LayoutControl1.TabIndex = 0
        Me.LayoutControl1.Text = "LayoutControl1"
        '
        'xtab1
        '
        Me.xtab1.Location = New System.Drawing.Point(7, 64)
        Me.xtab1.Name = "xtab1"
        Me.xtab1.SelectedTabPage = Me.XtraTabPage1
        Me.xtab1.Size = New System.Drawing.Size(948, 389)
        Me.xtab1.TabIndex = 10
        Me.xtab1.TabPages.AddRange(New DevExpress.XtraTab.XtraTabPage() {Me.XtraTabPage1, Me.XtraTabPage2, Me.XtraTabPage3, Me.XtraTabPage4, Me.XtraTabPage5})
        '
        'XtraTabPage1
        '
        Me.XtraTabPage1.Controls.Add(Me.LayoutControl2)
        Me.XtraTabPage1.Name = "XtraTabPage1"
        Me.XtraTabPage1.Size = New System.Drawing.Size(939, 359)
        Me.XtraTabPage1.Text = "Purchases"
        '
        'LayoutControl2
        '
        Me.LayoutControl2.Controls.Add(Me.gcPurchases)
        Me.LayoutControl2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LayoutControl2.Location = New System.Drawing.Point(0, 0)
        Me.LayoutControl2.Name = "LayoutControl2"
        Me.LayoutControl2.Root = Me.LayoutControlGroup2
        Me.LayoutControl2.Size = New System.Drawing.Size(939, 359)
        Me.LayoutControl2.TabIndex = 0
        Me.LayoutControl2.Text = "LayoutControl2"
        '
        'gcPurchases
        '
        Me.gcPurchases.Location = New System.Drawing.Point(7, 7)
        Me.gcPurchases.MainView = Me.gvPurchases
        Me.gcPurchases.Name = "gcPurchases"
        Me.gcPurchases.Size = New System.Drawing.Size(926, 346)
        Me.gcPurchases.TabIndex = 4
        Me.gcPurchases.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.gvPurchases})
        '
        'gvPurchases
        '
        Me.gvPurchases.Appearance.Row.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gvPurchases.Appearance.Row.Options.UseFont = True
        Me.gvPurchases.AppearancePrint.FooterPanel.BackColor = System.Drawing.Color.White
        Me.gvPurchases.AppearancePrint.FooterPanel.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gvPurchases.AppearancePrint.FooterPanel.Options.UseBackColor = True
        Me.gvPurchases.AppearancePrint.FooterPanel.Options.UseFont = True
        Me.gvPurchases.AppearancePrint.Row.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gvPurchases.AppearancePrint.Row.Options.UseFont = True
        Me.gvPurchases.Columns.AddRange(New DevExpress.XtraGrid.Columns.GridColumn() {Me.GridColumn1, Me.GridColumn2, Me.GridColumn3, Me.GridColumn4, Me.GridColumn5, Me.GridColumn6, Me.GridColumn7, Me.GridColumn8, Me.GridColumn9, Me.GridColumn10})
        Me.gvPurchases.GridControl = Me.gcPurchases
        Me.gvPurchases.GroupSummary.AddRange(New DevExpress.XtraGrid.GridSummaryItem() {New DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Count, "SupplierName", Me.GridColumn3, ""), New DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "Total", Me.GridColumn5, "{0:n2}"), New DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "Tax", Me.GridColumn6, "{0:n2}"), New DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "Discount", Me.GridColumn7, "{0:n2}"), New DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "GrandTotal", Me.GridColumn8, "{0:n2}"), New DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "AmountPaid", Me.GridColumn9, "{0:n2}"), New DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "Balance", Me.GridColumn10, "{0:n2}")})
        Me.gvPurchases.Name = "gvPurchases"
        Me.gvPurchases.OptionsBehavior.Editable = False
        Me.gvPurchases.OptionsPrint.UsePrintStyles = True
        Me.gvPurchases.OptionsView.ColumnAutoWidth = False
        Me.gvPurchases.OptionsView.ShowAutoFilterRow = True
        Me.gvPurchases.OptionsView.ShowFooter = True
        '
        'GridColumn1
        '
        Me.GridColumn1.Caption = "Purchase No"
        Me.GridColumn1.FieldName = "PurchaseNo"
        Me.GridColumn1.Name = "GridColumn1"
        Me.GridColumn1.SummaryItem.FieldName = "PurchaeNo"
        Me.GridColumn1.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Count
        Me.GridColumn1.Visible = True
        Me.GridColumn1.VisibleIndex = 0
        Me.GridColumn1.Width = 100
        '
        'GridColumn2
        '
        Me.GridColumn2.Caption = "Vendor No"
        Me.GridColumn2.FieldName = "SupplierNo"
        Me.GridColumn2.Name = "GridColumn2"
        Me.GridColumn2.Visible = True
        Me.GridColumn2.VisibleIndex = 1
        Me.GridColumn2.Width = 100
        '
        'GridColumn3
        '
        Me.GridColumn3.Caption = "Vendor Name"
        Me.GridColumn3.FieldName = "SupplierName"
        Me.GridColumn3.Name = "GridColumn3"
        Me.GridColumn3.Visible = True
        Me.GridColumn3.VisibleIndex = 2
        Me.GridColumn3.Width = 150
        '
        'GridColumn4
        '
        Me.GridColumn4.Caption = "Date"
        Me.GridColumn4.DisplayFormat.FormatString = "dd-MMM-yy"
        Me.GridColumn4.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.GridColumn4.FieldName = "PurchaseDate"
        Me.GridColumn4.Name = "GridColumn4"
        Me.GridColumn4.Visible = True
        Me.GridColumn4.VisibleIndex = 3
        Me.GridColumn4.Width = 100
        '
        'GridColumn5
        '
        Me.GridColumn5.Caption = "Total"
        Me.GridColumn5.DisplayFormat.FormatString = "{0:n2}"
        Me.GridColumn5.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.GridColumn5.FieldName = "Total"
        Me.GridColumn5.Name = "GridColumn5"
        Me.GridColumn5.SummaryItem.DisplayFormat = "{0:n2}"
        Me.GridColumn5.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum
        Me.GridColumn5.Width = 100
        '
        'GridColumn6
        '
        Me.GridColumn6.Caption = "Tax"
        Me.GridColumn6.DisplayFormat.FormatString = "{0:n2}"
        Me.GridColumn6.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.GridColumn6.FieldName = "Tax"
        Me.GridColumn6.Name = "GridColumn6"
        Me.GridColumn6.SummaryItem.DisplayFormat = "{0:n2}"
        Me.GridColumn6.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum
        Me.GridColumn6.Width = 100
        '
        'GridColumn7
        '
        Me.GridColumn7.Caption = "Discount"
        Me.GridColumn7.DisplayFormat.FormatString = "{0:n2}"
        Me.GridColumn7.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.GridColumn7.FieldName = "Discount"
        Me.GridColumn7.Name = "GridColumn7"
        Me.GridColumn7.SummaryItem.DisplayFormat = "{0:n2}"
        Me.GridColumn7.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum
        Me.GridColumn7.Width = 100
        '
        'GridColumn8
        '
        Me.GridColumn8.Caption = "Grand Total"
        Me.GridColumn8.DisplayFormat.FormatString = "{0:n2}"
        Me.GridColumn8.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.GridColumn8.FieldName = "GrandTotal"
        Me.GridColumn8.Name = "GridColumn8"
        Me.GridColumn8.SummaryItem.DisplayFormat = "{0:n2}"
        Me.GridColumn8.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum
        Me.GridColumn8.Visible = True
        Me.GridColumn8.VisibleIndex = 4
        Me.GridColumn8.Width = 100
        '
        'GridColumn9
        '
        Me.GridColumn9.Caption = "Amount Paid"
        Me.GridColumn9.DisplayFormat.FormatString = "{0:n2}"
        Me.GridColumn9.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.GridColumn9.FieldName = "AmountPaid"
        Me.GridColumn9.Name = "GridColumn9"
        Me.GridColumn9.SummaryItem.DisplayFormat = "{0:n2}"
        Me.GridColumn9.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum
        Me.GridColumn9.Visible = True
        Me.GridColumn9.VisibleIndex = 5
        Me.GridColumn9.Width = 100
        '
        'GridColumn10
        '
        Me.GridColumn10.Caption = "Balance"
        Me.GridColumn10.DisplayFormat.FormatString = "{0:n2}"
        Me.GridColumn10.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.GridColumn10.FieldName = "Balance"
        Me.GridColumn10.Name = "GridColumn10"
        Me.GridColumn10.SummaryItem.DisplayFormat = "{0:n2}"
        Me.GridColumn10.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum
        Me.GridColumn10.Visible = True
        Me.GridColumn10.VisibleIndex = 6
        Me.GridColumn10.Width = 100
        '
        'LayoutControlGroup2
        '
        Me.LayoutControlGroup2.CustomizationFormText = "LayoutControlGroup2"
        Me.LayoutControlGroup2.Items.AddRange(New DevExpress.XtraLayout.BaseLayoutItem() {Me.LayoutControlItem1})
        Me.LayoutControlGroup2.Location = New System.Drawing.Point(0, 0)
        Me.LayoutControlGroup2.Name = "LayoutControlGroup2"
        Me.LayoutControlGroup2.Size = New System.Drawing.Size(939, 359)
        Me.LayoutControlGroup2.Spacing = New DevExpress.XtraLayout.Utils.Padding(0, 0, 0, 0)
        Me.LayoutControlGroup2.Text = "LayoutControlGroup2"
        Me.LayoutControlGroup2.TextVisible = False
        '
        'LayoutControlItem1
        '
        Me.LayoutControlItem1.Control = Me.gcPurchases
        Me.LayoutControlItem1.CustomizationFormText = "LayoutControlItem1"
        Me.LayoutControlItem1.Location = New System.Drawing.Point(0, 0)
        Me.LayoutControlItem1.Name = "LayoutControlItem1"
        Me.LayoutControlItem1.Size = New System.Drawing.Size(937, 357)
        Me.LayoutControlItem1.Text = "LayoutControlItem1"
        Me.LayoutControlItem1.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem1.TextSize = New System.Drawing.Size(0, 0)
        Me.LayoutControlItem1.TextToControlDistance = 0
        Me.LayoutControlItem1.TextVisible = False
        '
        'XtraTabPage2
        '
        Me.XtraTabPage2.Controls.Add(Me.LayoutControl3)
        Me.XtraTabPage2.Name = "XtraTabPage2"
        Me.XtraTabPage2.Size = New System.Drawing.Size(939, 359)
        Me.XtraTabPage2.Text = "Sales"
        '
        'LayoutControl3
        '
        Me.LayoutControl3.Controls.Add(Me.gcSales)
        Me.LayoutControl3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LayoutControl3.Location = New System.Drawing.Point(0, 0)
        Me.LayoutControl3.Name = "LayoutControl3"
        Me.LayoutControl3.Root = Me.LayoutControlGroup3
        Me.LayoutControl3.Size = New System.Drawing.Size(939, 359)
        Me.LayoutControl3.TabIndex = 0
        Me.LayoutControl3.Text = "LayoutControl3"
        '
        'gcSales
        '
        Me.gcSales.Location = New System.Drawing.Point(7, 7)
        Me.gcSales.MainView = Me.gvSales
        Me.gcSales.Name = "gcSales"
        Me.gcSales.Size = New System.Drawing.Size(926, 346)
        Me.gcSales.TabIndex = 4
        Me.gcSales.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.gvSales})
        '
        'gvSales
        '
        Me.gvSales.Appearance.Row.BackColor = System.Drawing.Color.White
        Me.gvSales.Appearance.Row.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gvSales.Appearance.Row.Options.UseBackColor = True
        Me.gvSales.Appearance.Row.Options.UseFont = True
        Me.gvSales.AppearancePrint.FooterPanel.BackColor = System.Drawing.Color.White
        Me.gvSales.AppearancePrint.FooterPanel.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gvSales.AppearancePrint.FooterPanel.Options.UseBackColor = True
        Me.gvSales.AppearancePrint.FooterPanel.Options.UseFont = True
        Me.gvSales.AppearancePrint.Row.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gvSales.AppearancePrint.Row.Options.UseFont = True
        Me.gvSales.Columns.AddRange(New DevExpress.XtraGrid.Columns.GridColumn() {Me.GridColumn11, Me.GridColumn12, Me.GridColumn13, Me.GridColumn14, Me.GridColumn15, Me.GridColumn16, Me.GridColumn17, Me.GridColumn18, Me.GridColumn19, Me.GridColumn20, Me.GridColumn37, Me.GridColumn38})
        Me.gvSales.GridControl = Me.gcSales
        Me.gvSales.GroupSummary.AddRange(New DevExpress.XtraGrid.GridSummaryItem() {New DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Count, "CustomerName", Me.GridColumn13, ""), New DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "Total", Me.GridColumn15, "{0:n2}"), New DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "Tax", Me.GridColumn16, "{0:n2}"), New DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "Discount", Me.GridColumn17, "{0:n2}"), New DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "GrandTotal", Me.GridColumn18, "{0:n2}"), New DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "PaidAmount", Me.GridColumn19, "{0:n2}"), New DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "Balance", Me.GridColumn20, "{0:n2}")})
        Me.gvSales.Name = "gvSales"
        Me.gvSales.OptionsBehavior.Editable = False
        Me.gvSales.OptionsPrint.UsePrintStyles = True
        Me.gvSales.OptionsView.ShowAutoFilterRow = True
        Me.gvSales.OptionsView.ShowFooter = True
        '
        'GridColumn11
        '
        Me.GridColumn11.Caption = "Sales No"
        Me.GridColumn11.FieldName = "SalesNo"
        Me.GridColumn11.Name = "GridColumn11"
        Me.GridColumn11.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Count
        Me.GridColumn11.Visible = True
        Me.GridColumn11.VisibleIndex = 0
        '
        'GridColumn12
        '
        Me.GridColumn12.Caption = "Customer No"
        Me.GridColumn12.FieldName = "CustomerNo"
        Me.GridColumn12.Name = "GridColumn12"
        Me.GridColumn12.Visible = True
        Me.GridColumn12.VisibleIndex = 1
        '
        'GridColumn13
        '
        Me.GridColumn13.Caption = "Customer Name"
        Me.GridColumn13.FieldName = "CustomerName"
        Me.GridColumn13.Name = "GridColumn13"
        Me.GridColumn13.Visible = True
        Me.GridColumn13.VisibleIndex = 2
        '
        'GridColumn14
        '
        Me.GridColumn14.Caption = "Sales Date"
        Me.GridColumn14.DisplayFormat.FormatString = "dd-MMM-yy"
        Me.GridColumn14.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.GridColumn14.FieldName = "SalesDate"
        Me.GridColumn14.Name = "GridColumn14"
        Me.GridColumn14.Visible = True
        Me.GridColumn14.VisibleIndex = 3
        '
        'GridColumn15
        '
        Me.GridColumn15.Caption = "Vehicle No"
        Me.GridColumn15.FieldName = "ReferenceNo"
        Me.GridColumn15.Name = "GridColumn15"
        Me.GridColumn15.Visible = True
        Me.GridColumn15.VisibleIndex = 4
        '
        'GridColumn16
        '
        Me.GridColumn16.Caption = "Tax"
        Me.GridColumn16.DisplayFormat.FormatString = "{0:n2}"
        Me.GridColumn16.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.GridColumn16.FieldName = "Tax"
        Me.GridColumn16.Name = "GridColumn16"
        Me.GridColumn16.SummaryItem.DisplayFormat = "{0:n2}"
        Me.GridColumn16.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum
        '
        'GridColumn17
        '
        Me.GridColumn17.Caption = "Discount"
        Me.GridColumn17.DisplayFormat.FormatString = "{0:n2}"
        Me.GridColumn17.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.GridColumn17.FieldName = "Discount"
        Me.GridColumn17.Name = "GridColumn17"
        Me.GridColumn17.SummaryItem.DisplayFormat = "{0:n2}"
        Me.GridColumn17.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum
        '
        'GridColumn18
        '
        Me.GridColumn18.Caption = "Grand Total"
        Me.GridColumn18.DisplayFormat.FormatString = "{0:n2}"
        Me.GridColumn18.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.GridColumn18.FieldName = "GrandTotal"
        Me.GridColumn18.Name = "GridColumn18"
        Me.GridColumn18.SummaryItem.DisplayFormat = "{0:n2}"
        Me.GridColumn18.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum
        Me.GridColumn18.Visible = True
        Me.GridColumn18.VisibleIndex = 5
        '
        'GridColumn19
        '
        Me.GridColumn19.Caption = "Amount Paid"
        Me.GridColumn19.DisplayFormat.FormatString = "{0:n2}"
        Me.GridColumn19.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.GridColumn19.FieldName = "AmountPaid"
        Me.GridColumn19.Name = "GridColumn19"
        Me.GridColumn19.SummaryItem.DisplayFormat = "{0:n2}"
        Me.GridColumn19.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum
        Me.GridColumn19.Visible = True
        Me.GridColumn19.VisibleIndex = 6
        '
        'GridColumn20
        '
        Me.GridColumn20.Caption = "Balance"
        Me.GridColumn20.DisplayFormat.FormatString = "{0:n2}"
        Me.GridColumn20.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.GridColumn20.FieldName = "Balance"
        Me.GridColumn20.Name = "GridColumn20"
        Me.GridColumn20.SummaryItem.DisplayFormat = "{0:n2}"
        Me.GridColumn20.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum
        Me.GridColumn20.Visible = True
        Me.GridColumn20.VisibleIndex = 7
        '
        'GridColumn37
        '
        Me.GridColumn37.Caption = "Sevice1"
        Me.GridColumn37.FieldName = "ServiceEmployer1"
        Me.GridColumn37.Name = "GridColumn37"
        '
        'GridColumn38
        '
        Me.GridColumn38.Caption = "Service2"
        Me.GridColumn38.FieldName = "ServiceEmployer2"
        Me.GridColumn38.Name = "GridColumn38"
        '
        'LayoutControlGroup3
        '
        Me.LayoutControlGroup3.CustomizationFormText = "LayoutControlGroup3"
        Me.LayoutControlGroup3.Items.AddRange(New DevExpress.XtraLayout.BaseLayoutItem() {Me.LayoutControlItem6})
        Me.LayoutControlGroup3.Location = New System.Drawing.Point(0, 0)
        Me.LayoutControlGroup3.Name = "LayoutControlGroup3"
        Me.LayoutControlGroup3.Size = New System.Drawing.Size(939, 359)
        Me.LayoutControlGroup3.Spacing = New DevExpress.XtraLayout.Utils.Padding(0, 0, 0, 0)
        Me.LayoutControlGroup3.Text = "LayoutControlGroup3"
        Me.LayoutControlGroup3.TextVisible = False
        '
        'LayoutControlItem6
        '
        Me.LayoutControlItem6.Control = Me.gcSales
        Me.LayoutControlItem6.CustomizationFormText = "LayoutControlItem6"
        Me.LayoutControlItem6.Location = New System.Drawing.Point(0, 0)
        Me.LayoutControlItem6.Name = "LayoutControlItem6"
        Me.LayoutControlItem6.Size = New System.Drawing.Size(937, 357)
        Me.LayoutControlItem6.Text = "LayoutControlItem6"
        Me.LayoutControlItem6.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem6.TextSize = New System.Drawing.Size(0, 0)
        Me.LayoutControlItem6.TextToControlDistance = 0
        Me.LayoutControlItem6.TextVisible = False
        '
        'XtraTabPage3
        '
        Me.XtraTabPage3.Controls.Add(Me.LayoutControl4)
        Me.XtraTabPage3.Name = "XtraTabPage3"
        Me.XtraTabPage3.Size = New System.Drawing.Size(939, 359)
        Me.XtraTabPage3.Text = "Expenses"
        '
        'LayoutControl4
        '
        Me.LayoutControl4.Controls.Add(Me.gcExpenses)
        Me.LayoutControl4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LayoutControl4.Location = New System.Drawing.Point(0, 0)
        Me.LayoutControl4.Name = "LayoutControl4"
        Me.LayoutControl4.Root = Me.LayoutControlGroup5
        Me.LayoutControl4.Size = New System.Drawing.Size(939, 359)
        Me.LayoutControl4.TabIndex = 0
        Me.LayoutControl4.Text = "LayoutControl4"
        '
        'gcExpenses
        '
        Me.gcExpenses.Location = New System.Drawing.Point(7, 7)
        Me.gcExpenses.MainView = Me.gvExpenses
        Me.gcExpenses.Name = "gcExpenses"
        Me.gcExpenses.Size = New System.Drawing.Size(926, 346)
        Me.gcExpenses.TabIndex = 4
        Me.gcExpenses.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.gvExpenses})
        '
        'gvExpenses
        '
        Me.gvExpenses.Appearance.Row.BackColor = System.Drawing.Color.White
        Me.gvExpenses.Appearance.Row.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gvExpenses.Appearance.Row.Options.UseBackColor = True
        Me.gvExpenses.Appearance.Row.Options.UseFont = True
        Me.gvExpenses.AppearancePrint.FooterPanel.BackColor = System.Drawing.Color.White
        Me.gvExpenses.AppearancePrint.FooterPanel.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gvExpenses.AppearancePrint.FooterPanel.Options.UseBackColor = True
        Me.gvExpenses.AppearancePrint.FooterPanel.Options.UseFont = True
        Me.gvExpenses.AppearancePrint.Row.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gvExpenses.AppearancePrint.Row.Options.UseFont = True
        Me.gvExpenses.Columns.AddRange(New DevExpress.XtraGrid.Columns.GridColumn() {Me.GridColumn21, Me.GridColumn22, Me.GridColumn23, Me.GridColumn24})
        Me.gvExpenses.GridControl = Me.gcExpenses
        Me.gvExpenses.GroupSummary.AddRange(New DevExpress.XtraGrid.GridSummaryItem() {New DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "Amount", Me.GridColumn23, "{0:n2}"), New DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Count, "ExpenseDate", Me.GridColumn22, "")})
        Me.gvExpenses.Name = "gvExpenses"
        Me.gvExpenses.OptionsBehavior.Editable = False
        Me.gvExpenses.OptionsPrint.UsePrintStyles = True
        Me.gvExpenses.OptionsView.ShowAutoFilterRow = True
        Me.gvExpenses.OptionsView.ShowFooter = True
        '
        'GridColumn21
        '
        Me.GridColumn21.Caption = "Expense Type"
        Me.GridColumn21.FieldName = "Description"
        Me.GridColumn21.Name = "GridColumn21"
        Me.GridColumn21.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Count
        Me.GridColumn21.Visible = True
        Me.GridColumn21.VisibleIndex = 0
        '
        'GridColumn22
        '
        Me.GridColumn22.Caption = "Expense Date"
        Me.GridColumn22.DisplayFormat.FormatString = "dd-MMM-yy"
        Me.GridColumn22.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime
        Me.GridColumn22.FieldName = "ExpenseDate"
        Me.GridColumn22.Name = "GridColumn22"
        '
        'GridColumn23
        '
        Me.GridColumn23.Caption = "Amount"
        Me.GridColumn23.DisplayFormat.FormatString = "{0:n2}"
        Me.GridColumn23.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.GridColumn23.FieldName = "Amount"
        Me.GridColumn23.Name = "GridColumn23"
        Me.GridColumn23.SummaryItem.DisplayFormat = "{0:n2}"
        Me.GridColumn23.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum
        Me.GridColumn23.Visible = True
        Me.GridColumn23.VisibleIndex = 1
        '
        'GridColumn24
        '
        Me.GridColumn24.Caption = "Note"
        Me.GridColumn24.FieldName = "Note"
        Me.GridColumn24.Name = "GridColumn24"
        '
        'LayoutControlGroup5
        '
        Me.LayoutControlGroup5.CustomizationFormText = "LayoutControlGroup5"
        Me.LayoutControlGroup5.Items.AddRange(New DevExpress.XtraLayout.BaseLayoutItem() {Me.LayoutControlItem7})
        Me.LayoutControlGroup5.Location = New System.Drawing.Point(0, 0)
        Me.LayoutControlGroup5.Name = "LayoutControlGroup5"
        Me.LayoutControlGroup5.Size = New System.Drawing.Size(939, 359)
        Me.LayoutControlGroup5.Spacing = New DevExpress.XtraLayout.Utils.Padding(0, 0, 0, 0)
        Me.LayoutControlGroup5.Text = "LayoutControlGroup5"
        Me.LayoutControlGroup5.TextVisible = False
        '
        'LayoutControlItem7
        '
        Me.LayoutControlItem7.Control = Me.gcExpenses
        Me.LayoutControlItem7.CustomizationFormText = "LayoutControlItem7"
        Me.LayoutControlItem7.Location = New System.Drawing.Point(0, 0)
        Me.LayoutControlItem7.Name = "LayoutControlItem7"
        Me.LayoutControlItem7.Size = New System.Drawing.Size(937, 357)
        Me.LayoutControlItem7.Text = "LayoutControlItem7"
        Me.LayoutControlItem7.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem7.TextSize = New System.Drawing.Size(0, 0)
        Me.LayoutControlItem7.TextToControlDistance = 0
        Me.LayoutControlItem7.TextVisible = False
        '
        'XtraTabPage4
        '
        Me.XtraTabPage4.Controls.Add(Me.LayoutControl5)
        Me.XtraTabPage4.Name = "XtraTabPage4"
        Me.XtraTabPage4.Size = New System.Drawing.Size(939, 359)
        Me.XtraTabPage4.Text = "Stock Items"
        '
        'LayoutControl5
        '
        Me.LayoutControl5.Controls.Add(Me.gcStockItems)
        Me.LayoutControl5.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LayoutControl5.Location = New System.Drawing.Point(0, 0)
        Me.LayoutControl5.Name = "LayoutControl5"
        Me.LayoutControl5.Root = Me.LayoutControlGroup6
        Me.LayoutControl5.Size = New System.Drawing.Size(939, 359)
        Me.LayoutControl5.TabIndex = 0
        Me.LayoutControl5.Text = "LayoutControl5"
        '
        'gcStockItems
        '
        Me.gcStockItems.Location = New System.Drawing.Point(7, 7)
        Me.gcStockItems.MainView = Me.gvStockItems
        Me.gcStockItems.Name = "gcStockItems"
        Me.gcStockItems.Size = New System.Drawing.Size(926, 346)
        Me.gcStockItems.TabIndex = 4
        Me.gcStockItems.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.gvStockItems})
        '
        'gvStockItems
        '
        Me.gvStockItems.Appearance.Row.BackColor = System.Drawing.Color.White
        Me.gvStockItems.Appearance.Row.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gvStockItems.Appearance.Row.Options.UseBackColor = True
        Me.gvStockItems.Appearance.Row.Options.UseFont = True
        Me.gvStockItems.AppearancePrint.FooterPanel.BackColor = System.Drawing.Color.White
        Me.gvStockItems.AppearancePrint.FooterPanel.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gvStockItems.AppearancePrint.FooterPanel.Options.UseBackColor = True
        Me.gvStockItems.AppearancePrint.FooterPanel.Options.UseFont = True
        Me.gvStockItems.AppearancePrint.Row.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gvStockItems.AppearancePrint.Row.Options.UseFont = True
        Me.gvStockItems.Columns.AddRange(New DevExpress.XtraGrid.Columns.GridColumn() {Me.GridColumn25, Me.GridColumn26, Me.GridColumn29, Me.GridColumn30})
        Me.gvStockItems.GridControl = Me.gcStockItems
        Me.gvStockItems.GroupSummary.AddRange(New DevExpress.XtraGrid.GridSummaryItem() {New DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Count, "Description", Me.GridColumn26, ""), New DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "Quantity", Me.GridColumn29, "{0:n2}"), New DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "Amount", Me.GridColumn30, "{0:n2}")})
        Me.gvStockItems.Name = "gvStockItems"
        Me.gvStockItems.OptionsBehavior.Editable = False
        Me.gvStockItems.OptionsPrint.UsePrintStyles = True
        Me.gvStockItems.OptionsView.ColumnAutoWidth = False
        Me.gvStockItems.OptionsView.ShowAutoFilterRow = True
        Me.gvStockItems.OptionsView.ShowFooter = True
        '
        'GridColumn25
        '
        Me.GridColumn25.Caption = "Stock Code"
        Me.GridColumn25.FieldName = "StockCode"
        Me.GridColumn25.Name = "GridColumn25"
        Me.GridColumn25.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Count
        Me.GridColumn25.Visible = True
        Me.GridColumn25.VisibleIndex = 0
        Me.GridColumn25.Width = 100
        '
        'GridColumn26
        '
        Me.GridColumn26.Caption = "Description"
        Me.GridColumn26.FieldName = "Description"
        Me.GridColumn26.Name = "GridColumn26"
        Me.GridColumn26.Visible = True
        Me.GridColumn26.VisibleIndex = 1
        Me.GridColumn26.Width = 150
        '
        'GridColumn29
        '
        Me.GridColumn29.Caption = "Quantity"
        Me.GridColumn29.FieldName = "Quantity"
        Me.GridColumn29.Name = "GridColumn29"
        Me.GridColumn29.Visible = True
        Me.GridColumn29.VisibleIndex = 2
        Me.GridColumn29.Width = 120
        '
        'GridColumn30
        '
        Me.GridColumn30.Caption = "Amount"
        Me.GridColumn30.DisplayFormat.FormatString = "{0:n2}"
        Me.GridColumn30.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.GridColumn30.FieldName = "Amount"
        Me.GridColumn30.Name = "GridColumn30"
        Me.GridColumn30.SummaryItem.DisplayFormat = "{0:n2}"
        Me.GridColumn30.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum
        Me.GridColumn30.Visible = True
        Me.GridColumn30.VisibleIndex = 3
        Me.GridColumn30.Width = 120
        '
        'LayoutControlGroup6
        '
        Me.LayoutControlGroup6.CustomizationFormText = "LayoutControlGroup6"
        Me.LayoutControlGroup6.Items.AddRange(New DevExpress.XtraLayout.BaseLayoutItem() {Me.LayoutControlItem8})
        Me.LayoutControlGroup6.Location = New System.Drawing.Point(0, 0)
        Me.LayoutControlGroup6.Name = "LayoutControlGroup6"
        Me.LayoutControlGroup6.Size = New System.Drawing.Size(939, 359)
        Me.LayoutControlGroup6.Spacing = New DevExpress.XtraLayout.Utils.Padding(0, 0, 0, 0)
        Me.LayoutControlGroup6.Text = "LayoutControlGroup6"
        Me.LayoutControlGroup6.TextVisible = False
        '
        'LayoutControlItem8
        '
        Me.LayoutControlItem8.Control = Me.gcStockItems
        Me.LayoutControlItem8.CustomizationFormText = "LayoutControlItem8"
        Me.LayoutControlItem8.Location = New System.Drawing.Point(0, 0)
        Me.LayoutControlItem8.Name = "LayoutControlItem8"
        Me.LayoutControlItem8.Size = New System.Drawing.Size(937, 357)
        Me.LayoutControlItem8.Text = "LayoutControlItem8"
        Me.LayoutControlItem8.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem8.TextSize = New System.Drawing.Size(0, 0)
        Me.LayoutControlItem8.TextToControlDistance = 0
        Me.LayoutControlItem8.TextVisible = False
        '
        'XtraTabPage5
        '
        Me.XtraTabPage5.Controls.Add(Me.LayoutControl6)
        Me.XtraTabPage5.Name = "XtraTabPage5"
        Me.XtraTabPage5.Size = New System.Drawing.Size(939, 359)
        Me.XtraTabPage5.Text = "Services"
        '
        'LayoutControl6
        '
        Me.LayoutControl6.Controls.Add(Me.gcServices)
        Me.LayoutControl6.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LayoutControl6.Location = New System.Drawing.Point(0, 0)
        Me.LayoutControl6.Name = "LayoutControl6"
        Me.LayoutControl6.Root = Me.LayoutControlGroup7
        Me.LayoutControl6.Size = New System.Drawing.Size(939, 359)
        Me.LayoutControl6.TabIndex = 0
        Me.LayoutControl6.Text = "LayoutControl6"
        '
        'gcServices
        '
        Me.gcServices.Location = New System.Drawing.Point(7, 7)
        Me.gcServices.MainView = Me.gvServices
        Me.gcServices.Name = "gcServices"
        Me.gcServices.Size = New System.Drawing.Size(926, 346)
        Me.gcServices.TabIndex = 4
        Me.gcServices.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.gvServices})
        '
        'gvServices
        '
        Me.gvServices.Appearance.Row.BackColor = System.Drawing.Color.White
        Me.gvServices.Appearance.Row.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gvServices.Appearance.Row.Options.UseBackColor = True
        Me.gvServices.Appearance.Row.Options.UseFont = True
        Me.gvServices.AppearancePrint.FooterPanel.BackColor = System.Drawing.Color.White
        Me.gvServices.AppearancePrint.FooterPanel.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gvServices.AppearancePrint.FooterPanel.Options.UseBackColor = True
        Me.gvServices.AppearancePrint.FooterPanel.Options.UseFont = True
        Me.gvServices.AppearancePrint.Row.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gvServices.AppearancePrint.Row.Options.UseFont = True
        Me.gvServices.Columns.AddRange(New DevExpress.XtraGrid.Columns.GridColumn() {Me.GridColumn31, Me.GridColumn32, Me.GridColumn35, Me.GridColumn36})
        Me.gvServices.GridControl = Me.gcServices
        Me.gvServices.GroupSummary.AddRange(New DevExpress.XtraGrid.GridSummaryItem() {New DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Count, "Description", Me.GridColumn32, ""), New DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "Quantity", Me.GridColumn35, ""), New DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "Amount", Me.GridColumn36, "{0:n2}")})
        Me.gvServices.Name = "gvServices"
        Me.gvServices.OptionsPrint.UsePrintStyles = True
        Me.gvServices.OptionsView.ColumnAutoWidth = False
        Me.gvServices.OptionsView.ShowAutoFilterRow = True
        Me.gvServices.OptionsView.ShowFooter = True
        '
        'GridColumn31
        '
        Me.GridColumn31.Caption = "Service Code"
        Me.GridColumn31.FieldName = "StockCode"
        Me.GridColumn31.Name = "GridColumn31"
        Me.GridColumn31.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Count
        Me.GridColumn31.Visible = True
        Me.GridColumn31.VisibleIndex = 0
        Me.GridColumn31.Width = 120
        '
        'GridColumn32
        '
        Me.GridColumn32.Caption = "Description"
        Me.GridColumn32.FieldName = "Description"
        Me.GridColumn32.Name = "GridColumn32"
        Me.GridColumn32.Visible = True
        Me.GridColumn32.VisibleIndex = 1
        Me.GridColumn32.Width = 150
        '
        'GridColumn35
        '
        Me.GridColumn35.Caption = "Quantity"
        Me.GridColumn35.FieldName = "Quantity"
        Me.GridColumn35.Name = "GridColumn35"
        Me.GridColumn35.Visible = True
        Me.GridColumn35.VisibleIndex = 2
        Me.GridColumn35.Width = 120
        '
        'GridColumn36
        '
        Me.GridColumn36.Caption = "Amount"
        Me.GridColumn36.DisplayFormat.FormatString = "{0:n2}"
        Me.GridColumn36.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.GridColumn36.FieldName = "Amount"
        Me.GridColumn36.Name = "GridColumn36"
        Me.GridColumn36.SummaryItem.DisplayFormat = "{0:n2}"
        Me.GridColumn36.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum
        Me.GridColumn36.Visible = True
        Me.GridColumn36.VisibleIndex = 3
        Me.GridColumn36.Width = 120
        '
        'LayoutControlGroup7
        '
        Me.LayoutControlGroup7.CustomizationFormText = "LayoutControlGroup7"
        Me.LayoutControlGroup7.Items.AddRange(New DevExpress.XtraLayout.BaseLayoutItem() {Me.LayoutControlItem9})
        Me.LayoutControlGroup7.Location = New System.Drawing.Point(0, 0)
        Me.LayoutControlGroup7.Name = "LayoutControlGroup7"
        Me.LayoutControlGroup7.Size = New System.Drawing.Size(939, 359)
        Me.LayoutControlGroup7.Spacing = New DevExpress.XtraLayout.Utils.Padding(0, 0, 0, 0)
        Me.LayoutControlGroup7.Text = "LayoutControlGroup7"
        Me.LayoutControlGroup7.TextVisible = False
        '
        'LayoutControlItem9
        '
        Me.LayoutControlItem9.Control = Me.gcServices
        Me.LayoutControlItem9.CustomizationFormText = "LayoutControlItem9"
        Me.LayoutControlItem9.Location = New System.Drawing.Point(0, 0)
        Me.LayoutControlItem9.Name = "LayoutControlItem9"
        Me.LayoutControlItem9.Size = New System.Drawing.Size(937, 357)
        Me.LayoutControlItem9.Text = "LayoutControlItem9"
        Me.LayoutControlItem9.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem9.TextSize = New System.Drawing.Size(0, 0)
        Me.LayoutControlItem9.TextToControlDistance = 0
        Me.LayoutControlItem9.TextVisible = False
        '
        'sbPrint
        '
        Me.sbPrint.Location = New System.Drawing.Point(763, 461)
        Me.sbPrint.Name = "sbPrint"
        Me.sbPrint.Size = New System.Drawing.Size(195, 30)
        Me.sbPrint.StyleController = Me.LayoutControl1
        Me.sbPrint.TabIndex = 9
        Me.sbPrint.Text = "Print"
        '
        'sbProcess
        '
        Me.sbProcess.Location = New System.Drawing.Point(622, 25)
        Me.sbProcess.Name = "sbProcess"
        Me.sbProcess.Size = New System.Drawing.Size(89, 28)
        Me.sbProcess.StyleController = Me.LayoutControl1
        Me.sbProcess.TabIndex = 2
        Me.sbProcess.Text = "Process"
        '
        'deToDate
        '
        Me.deToDate.EditValue = Nothing
        Me.deToDate.EnterMoveNextControl = True
        Me.deToDate.Location = New System.Drawing.Point(454, 28)
        Me.deToDate.Name = "deToDate"
        Me.deToDate.Properties.AllowNullInput = DevExpress.Utils.DefaultBoolean.[False]
        Me.deToDate.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.deToDate.Properties.DisplayFormat.FormatString = "dd-MMM-yy"
        Me.deToDate.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime
        Me.deToDate.Properties.EditFormat.FormatString = "dd-MMM-yy"
        Me.deToDate.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.DateTime
        Me.deToDate.Properties.Mask.EditMask = "dd-MMM-yy"
        Me.deToDate.Properties.VistaTimeProperties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton})
        Me.deToDate.Size = New System.Drawing.Size(160, 20)
        Me.deToDate.StyleController = Me.LayoutControl1
        Me.deToDate.TabIndex = 1
        CompareAgainstControlValidationRule1.CompareControlOperator = DevExpress.XtraEditors.DXErrorProvider.CompareControlOperator.GreaterOrEqual
        CompareAgainstControlValidationRule1.Control = Me.deFromDate
        CompareAgainstControlValidationRule1.ErrorText = "Value must be a recent Date"
        Me.dxvpSalesSummary.SetValidationRule(Me.deToDate, CompareAgainstControlValidationRule1)
        '
        'LayoutControlGroup1
        '
        Me.LayoutControlGroup1.CustomizationFormText = "LayoutControlGroup1"
        Me.LayoutControlGroup1.Items.AddRange(New DevExpress.XtraLayout.BaseLayoutItem() {Me.EmptySpaceItem1, Me.EmptySpaceItem2, Me.LayoutControlGroup4, Me.LayoutControlItem2, Me.EmptySpaceItem4, Me.LayoutControlItem3})
        Me.LayoutControlGroup1.Location = New System.Drawing.Point(0, 0)
        Me.LayoutControlGroup1.Name = "LayoutControlGroup1"
        Me.LayoutControlGroup1.Size = New System.Drawing.Size(961, 494)
        Me.LayoutControlGroup1.Spacing = New DevExpress.XtraLayout.Utils.Padding(0, 0, 0, 0)
        Me.LayoutControlGroup1.Text = "LayoutControlGroup1"
        Me.LayoutControlGroup1.TextVisible = False
        '
        'EmptySpaceItem1
        '
        Me.EmptySpaceItem1.CustomizationFormText = "EmptySpaceItem1"
        Me.EmptySpaceItem1.Location = New System.Drawing.Point(0, 0)
        Me.EmptySpaceItem1.Name = "EmptySpaceItem1"
        Me.EmptySpaceItem1.Size = New System.Drawing.Size(215, 57)
        Me.EmptySpaceItem1.Text = "EmptySpaceItem1"
        Me.EmptySpaceItem1.TextSize = New System.Drawing.Size(0, 0)
        '
        'EmptySpaceItem2
        '
        Me.EmptySpaceItem2.CustomizationFormText = "EmptySpaceItem2"
        Me.EmptySpaceItem2.Location = New System.Drawing.Point(715, 0)
        Me.EmptySpaceItem2.Name = "EmptySpaceItem2"
        Me.EmptySpaceItem2.Size = New System.Drawing.Size(244, 57)
        Me.EmptySpaceItem2.Text = "EmptySpaceItem2"
        Me.EmptySpaceItem2.TextSize = New System.Drawing.Size(0, 0)
        '
        'LayoutControlGroup4
        '
        Me.LayoutControlGroup4.CustomizationFormText = "Select Date Range to extract Cash Collection"
        Me.LayoutControlGroup4.Items.AddRange(New DevExpress.XtraLayout.BaseLayoutItem() {Me.From, Me.LayoutControlItem4, Me.LayoutControlItem5})
        Me.LayoutControlGroup4.Location = New System.Drawing.Point(215, 0)
        Me.LayoutControlGroup4.Name = "LayoutControlGroup4"
        Me.LayoutControlGroup4.Size = New System.Drawing.Size(500, 57)
        Me.LayoutControlGroup4.Text = "Select Date Range to extract Daily Collection"
        '
        'From
        '
        Me.From.Control = Me.deFromDate
        Me.From.CustomizationFormText = "From"
        Me.From.Location = New System.Drawing.Point(0, 0)
        Me.From.Name = "From"
        Me.From.Size = New System.Drawing.Size(200, 33)
        Me.From.Text = "From"
        Me.From.TextLocation = DevExpress.Utils.Locations.Left
        Me.From.TextSize = New System.Drawing.Size(24, 20)
        '
        'LayoutControlItem4
        '
        Me.LayoutControlItem4.Control = Me.deToDate
        Me.LayoutControlItem4.CustomizationFormText = "To"
        Me.LayoutControlItem4.Location = New System.Drawing.Point(200, 0)
        Me.LayoutControlItem4.Name = "LayoutControlItem4"
        Me.LayoutControlItem4.Size = New System.Drawing.Size(200, 33)
        Me.LayoutControlItem4.Text = "To"
        Me.LayoutControlItem4.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem4.TextSize = New System.Drawing.Size(24, 20)
        '
        'LayoutControlItem5
        '
        Me.LayoutControlItem5.Control = Me.sbProcess
        Me.LayoutControlItem5.CustomizationFormText = "LayoutControlItem5"
        Me.LayoutControlItem5.Location = New System.Drawing.Point(400, 0)
        Me.LayoutControlItem5.MaxSize = New System.Drawing.Size(0, 33)
        Me.LayoutControlItem5.MinSize = New System.Drawing.Size(61, 33)
        Me.LayoutControlItem5.Name = "LayoutControlItem5"
        Me.LayoutControlItem5.Padding = New DevExpress.XtraLayout.Utils.Padding(2, 2, 2, 2)
        Me.LayoutControlItem5.Size = New System.Drawing.Size(94, 33)
        Me.LayoutControlItem5.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom
        Me.LayoutControlItem5.Text = "LayoutControlItem5"
        Me.LayoutControlItem5.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem5.TextSize = New System.Drawing.Size(0, 0)
        Me.LayoutControlItem5.TextToControlDistance = 0
        Me.LayoutControlItem5.TextVisible = False
        '
        'LayoutControlItem2
        '
        Me.LayoutControlItem2.Control = Me.sbPrint
        Me.LayoutControlItem2.CustomizationFormText = "LayoutControlItem2"
        Me.LayoutControlItem2.Location = New System.Drawing.Point(759, 457)
        Me.LayoutControlItem2.MaxSize = New System.Drawing.Size(200, 35)
        Me.LayoutControlItem2.MinSize = New System.Drawing.Size(200, 35)
        Me.LayoutControlItem2.Name = "LayoutControlItem2"
        Me.LayoutControlItem2.Padding = New DevExpress.XtraLayout.Utils.Padding(2, 2, 2, 2)
        Me.LayoutControlItem2.Size = New System.Drawing.Size(200, 35)
        Me.LayoutControlItem2.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom
        Me.LayoutControlItem2.Text = "LayoutControlItem2"
        Me.LayoutControlItem2.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem2.TextSize = New System.Drawing.Size(0, 0)
        Me.LayoutControlItem2.TextToControlDistance = 0
        Me.LayoutControlItem2.TextVisible = False
        '
        'EmptySpaceItem4
        '
        Me.EmptySpaceItem4.CustomizationFormText = "EmptySpaceItem4"
        Me.EmptySpaceItem4.Location = New System.Drawing.Point(0, 457)
        Me.EmptySpaceItem4.Name = "EmptySpaceItem4"
        Me.EmptySpaceItem4.Size = New System.Drawing.Size(759, 35)
        Me.EmptySpaceItem4.Text = "EmptySpaceItem4"
        Me.EmptySpaceItem4.TextSize = New System.Drawing.Size(0, 0)
        '
        'LayoutControlItem3
        '
        Me.LayoutControlItem3.Control = Me.xtab1
        Me.LayoutControlItem3.CustomizationFormText = "LayoutControlItem3"
        Me.LayoutControlItem3.Location = New System.Drawing.Point(0, 57)
        Me.LayoutControlItem3.Name = "LayoutControlItem3"
        Me.LayoutControlItem3.Size = New System.Drawing.Size(959, 400)
        Me.LayoutControlItem3.Text = "LayoutControlItem3"
        Me.LayoutControlItem3.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem3.TextSize = New System.Drawing.Size(0, 0)
        Me.LayoutControlItem3.TextToControlDistance = 0
        Me.LayoutControlItem3.TextVisible = False
        '
        'frmDailySummary
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(961, 494)
        Me.Controls.Add(Me.LayoutControl1)
        Me.KeyPreview = True
        Me.Name = "frmDailySummary"
        Me.Text = "Daily Summary"
        CType(Me.deFromDate.Properties.VistaTimeProperties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.deFromDate.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.LayoutControl1.ResumeLayout(False)
        CType(Me.xtab1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.xtab1.ResumeLayout(False)
        Me.XtraTabPage1.ResumeLayout(False)
        CType(Me.LayoutControl2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.LayoutControl2.ResumeLayout(False)
        CType(Me.gcPurchases, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gvPurchases, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlGroup2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.XtraTabPage2.ResumeLayout(False)
        CType(Me.LayoutControl3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.LayoutControl3.ResumeLayout(False)
        CType(Me.gcSales, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gvSales, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlGroup3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem6, System.ComponentModel.ISupportInitialize).EndInit()
        Me.XtraTabPage3.ResumeLayout(False)
        CType(Me.LayoutControl4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.LayoutControl4.ResumeLayout(False)
        CType(Me.gcExpenses, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gvExpenses, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlGroup5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem7, System.ComponentModel.ISupportInitialize).EndInit()
        Me.XtraTabPage4.ResumeLayout(False)
        CType(Me.LayoutControl5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.LayoutControl5.ResumeLayout(False)
        CType(Me.gcStockItems, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gvStockItems, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlGroup6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem8, System.ComponentModel.ISupportInitialize).EndInit()
        Me.XtraTabPage5.ResumeLayout(False)
        CType(Me.LayoutControl6, System.ComponentModel.ISupportInitialize).EndInit()
        Me.LayoutControl6.ResumeLayout(False)
        CType(Me.gcServices, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gvServices, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlGroup7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem9, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.deToDate.Properties.VistaTimeProperties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.deToDate.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlGroup1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.EmptySpaceItem1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.EmptySpaceItem2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlGroup4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.From, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.EmptySpaceItem4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dxvpSalesSummary, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents LayoutControl1 As DevExpress.XtraLayout.LayoutControl
    Friend WithEvents LayoutControlGroup1 As DevExpress.XtraLayout.LayoutControlGroup
    Friend WithEvents deFromDate As DevExpress.XtraEditors.DateEdit
    Friend WithEvents From As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents deToDate As DevExpress.XtraEditors.DateEdit
    Friend WithEvents LayoutControlItem4 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents EmptySpaceItem1 As DevExpress.XtraLayout.EmptySpaceItem
    Friend WithEvents EmptySpaceItem2 As DevExpress.XtraLayout.EmptySpaceItem
    Friend WithEvents LayoutControlGroup4 As DevExpress.XtraLayout.LayoutControlGroup
    Friend WithEvents sbProcess As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents LayoutControlItem5 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents dxvpSalesSummary As DevExpress.XtraEditors.DXErrorProvider.DXValidationProvider
    Friend WithEvents sbPrint As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents LayoutControlItem2 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents EmptySpaceItem4 As DevExpress.XtraLayout.EmptySpaceItem
    Friend WithEvents xtab1 As DevExpress.XtraTab.XtraTabControl
    Friend WithEvents XtraTabPage1 As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents XtraTabPage2 As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents XtraTabPage3 As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents XtraTabPage4 As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents XtraTabPage5 As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents LayoutControlItem3 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents LayoutControl2 As DevExpress.XtraLayout.LayoutControl
    Friend WithEvents gcPurchases As DevExpress.XtraGrid.GridControl
    Friend WithEvents gvPurchases As DevExpress.XtraGrid.Views.Grid.GridView
    Friend WithEvents LayoutControlGroup2 As DevExpress.XtraLayout.LayoutControlGroup
    Friend WithEvents LayoutControlItem1 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents GridColumn1 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn2 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn3 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn4 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn5 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn6 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn7 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn8 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn9 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn10 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents LayoutControl3 As DevExpress.XtraLayout.LayoutControl
    Friend WithEvents gcSales As DevExpress.XtraGrid.GridControl
    Friend WithEvents gvSales As DevExpress.XtraGrid.Views.Grid.GridView
    Friend WithEvents LayoutControlGroup3 As DevExpress.XtraLayout.LayoutControlGroup
    Friend WithEvents LayoutControlItem6 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents GridColumn11 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn12 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn13 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn14 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn15 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn16 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn17 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn18 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn19 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn20 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents LayoutControl4 As DevExpress.XtraLayout.LayoutControl
    Friend WithEvents LayoutControlGroup5 As DevExpress.XtraLayout.LayoutControlGroup
    Friend WithEvents gcExpenses As DevExpress.XtraGrid.GridControl
    Friend WithEvents gvExpenses As DevExpress.XtraGrid.Views.Grid.GridView
    Friend WithEvents LayoutControlItem7 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents GridColumn21 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn22 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn23 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn24 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents LayoutControl5 As DevExpress.XtraLayout.LayoutControl
    Friend WithEvents gcStockItems As DevExpress.XtraGrid.GridControl
    Friend WithEvents gvStockItems As DevExpress.XtraGrid.Views.Grid.GridView
    Friend WithEvents LayoutControlGroup6 As DevExpress.XtraLayout.LayoutControlGroup
    Friend WithEvents LayoutControlItem8 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents GridColumn25 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn26 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn29 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn30 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents LayoutControl6 As DevExpress.XtraLayout.LayoutControl
    Friend WithEvents gcServices As DevExpress.XtraGrid.GridControl
    Friend WithEvents gvServices As DevExpress.XtraGrid.Views.Grid.GridView
    Friend WithEvents LayoutControlGroup7 As DevExpress.XtraLayout.LayoutControlGroup
    Friend WithEvents LayoutControlItem9 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents GridColumn31 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn32 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn35 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn36 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn37 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn38 As DevExpress.XtraGrid.Columns.GridColumn
End Class
