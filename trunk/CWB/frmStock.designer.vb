<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmStock
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
        Dim ConditionValidationRule1 As DevExpress.XtraEditors.DXErrorProvider.ConditionValidationRule = New DevExpress.XtraEditors.DXErrorProvider.ConditionValidationRule
        Dim ConditionValidationRule2 As DevExpress.XtraEditors.DXErrorProvider.ConditionValidationRule = New DevExpress.XtraEditors.DXErrorProvider.ConditionValidationRule
        Dim ConditionValidationRule3 As DevExpress.XtraEditors.DXErrorProvider.ConditionValidationRule = New DevExpress.XtraEditors.DXErrorProvider.ConditionValidationRule
        Dim ConditionValidatonRule1 As DevExpress.XtraEditors.DXErrorProvider.ConditionValidatonRule = New DevExpress.XtraEditors.DXErrorProvider.ConditionValidatonRule
        Dim ConditionValidatonRule2 As DevExpress.XtraEditors.DXErrorProvider.ConditionValidatonRule = New DevExpress.XtraEditors.DXErrorProvider.ConditionValidatonRule
        Me.BarManager1 = New DevExpress.XtraBars.BarManager(Me.components)
        Me.Bar2 = New DevExpress.XtraBars.Bar
        Me.bbSave = New DevExpress.XtraBars.BarButtonItem
        Me.bbNew = New DevExpress.XtraBars.BarButtonItem
        Me.bbDelete = New DevExpress.XtraBars.BarButtonItem
        Me.bbDisplaySelected = New DevExpress.XtraBars.BarButtonItem
        Me.bbRefresh = New DevExpress.XtraBars.BarButtonItem
        Me.bbPrint = New DevExpress.XtraBars.BarButtonItem
        Me.bbClose = New DevExpress.XtraBars.BarButtonItem
        Me.barDockControlTop = New DevExpress.XtraBars.BarDockControl
        Me.barDockControlBottom = New DevExpress.XtraBars.BarDockControl
        Me.barDockControlLeft = New DevExpress.XtraBars.BarDockControl
        Me.barDockControlRight = New DevExpress.XtraBars.BarDockControl
        Me.bbShowAll = New DevExpress.XtraBars.BarButtonItem
        Me.xTab1 = New DevExpress.XtraTab.XtraTabControl
        Me.XtraTabPage1 = New DevExpress.XtraTab.XtraTabPage
        Me.LayoutControl1 = New DevExpress.XtraLayout.LayoutControl
        Me.ceStockType = New DevExpress.XtraEditors.ComboBoxEdit
        Me.seCurrentStockValue = New DevExpress.XtraEditors.SpinEdit
        Me.seCurrentStockBalance = New DevExpress.XtraEditors.SpinEdit
        Me.lblID = New DevExpress.XtraEditors.LabelControl
        Me.seOpeningStockValue = New DevExpress.XtraEditors.SpinEdit
        Me.seOpeningStockBalance = New DevExpress.XtraEditors.SpinEdit
        Me.seReorderLevel = New DevExpress.XtraEditors.SpinEdit
        Me.lupSupplier = New DevExpress.XtraEditors.LookUpEdit
        Me.seSalesPrice = New DevExpress.XtraEditors.SpinEdit
        Me.sePurchasePrice = New DevExpress.XtraEditors.SpinEdit
        Me.LupStockCategory = New DevExpress.XtraEditors.LookUpEdit
        Me.teDescription = New DevExpress.XtraEditors.TextEdit
        Me.teStockCode = New DevExpress.XtraEditors.TextEdit
        Me.LayoutControlGroup1 = New DevExpress.XtraLayout.LayoutControlGroup
        Me.LayoutControlItem1 = New DevExpress.XtraLayout.LayoutControlItem
        Me.LayoutControlItem2 = New DevExpress.XtraLayout.LayoutControlItem
        Me.LayoutControlItem10 = New DevExpress.XtraLayout.LayoutControlItem
        Me.EmptySpaceItem1 = New DevExpress.XtraLayout.EmptySpaceItem
        Me.EmptySpaceItem4 = New DevExpress.XtraLayout.EmptySpaceItem
        Me.LayoutControlItem3 = New DevExpress.XtraLayout.LayoutControlItem
        Me.EmptySpaceItem3 = New DevExpress.XtraLayout.EmptySpaceItem
        Me.LayoutControlItem13 = New DevExpress.XtraLayout.LayoutControlItem
        Me.LayoutControlItem4 = New DevExpress.XtraLayout.LayoutControlItem
        Me.LayoutControlItem5 = New DevExpress.XtraLayout.LayoutControlItem
        Me.LayoutControlItem6 = New DevExpress.XtraLayout.LayoutControlItem
        Me.LayoutControlItem9 = New DevExpress.XtraLayout.LayoutControlItem
        Me.LayoutControlItem8 = New DevExpress.XtraLayout.LayoutControlItem
        Me.LayoutControlItem7 = New DevExpress.XtraLayout.LayoutControlItem
        Me.EmptySpaceItem2 = New DevExpress.XtraLayout.EmptySpaceItem
        Me.EmptySpaceItem5 = New DevExpress.XtraLayout.EmptySpaceItem
        Me.EmptySpaceItem6 = New DevExpress.XtraLayout.EmptySpaceItem
        Me.EmptySpaceItem7 = New DevExpress.XtraLayout.EmptySpaceItem
        Me.EmptySpaceItem8 = New DevExpress.XtraLayout.EmptySpaceItem
        Me.LayoutControlGroup2 = New DevExpress.XtraLayout.LayoutControlGroup
        Me.LayoutControlItem11 = New DevExpress.XtraLayout.LayoutControlItem
        Me.LayoutControlItem12 = New DevExpress.XtraLayout.LayoutControlItem
        Me.EmptySpaceItem9 = New DevExpress.XtraLayout.EmptySpaceItem
        Me.XtraTabPage2 = New DevExpress.XtraTab.XtraTabPage
        Me.gcStockItems = New DevExpress.XtraGrid.GridControl
        Me.gvStockItems = New DevExpress.XtraGrid.Views.Grid.GridView
        Me.GridColumn1 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn10 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn2 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn17 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn3 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn4 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn5 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn6 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn7 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn8 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn9 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn11 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn12 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn13 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn14 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn15 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn16 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.dxvpStockItems = New DevExpress.XtraEditors.DXErrorProvider.DXValidationProvider(Me.components)
        CType(Me.BarManager1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.xTab1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.xTab1.SuspendLayout()
        Me.XtraTabPage1.SuspendLayout()
        CType(Me.LayoutControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.LayoutControl1.SuspendLayout()
        CType(Me.ceStockType.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.seCurrentStockValue.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.seCurrentStockBalance.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.seOpeningStockValue.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.seOpeningStockBalance.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.seReorderLevel.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lupSupplier.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.seSalesPrice.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.sePurchasePrice.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LupStockCategory.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.teDescription.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.teStockCode.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlGroup1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem10, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.EmptySpaceItem1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.EmptySpaceItem4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.EmptySpaceItem3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem13, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem9, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem8, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.EmptySpaceItem2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.EmptySpaceItem5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.EmptySpaceItem6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.EmptySpaceItem7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.EmptySpaceItem8, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlGroup2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem11, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem12, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.EmptySpaceItem9, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.XtraTabPage2.SuspendLayout()
        CType(Me.gcStockItems, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gvStockItems, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dxvpStockItems, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'BarManager1
        '
        Me.BarManager1.Bars.AddRange(New DevExpress.XtraBars.Bar() {Me.Bar2})
        Me.BarManager1.DockControls.Add(Me.barDockControlTop)
        Me.BarManager1.DockControls.Add(Me.barDockControlBottom)
        Me.BarManager1.DockControls.Add(Me.barDockControlLeft)
        Me.BarManager1.DockControls.Add(Me.barDockControlRight)
        Me.BarManager1.Form = Me
        Me.BarManager1.Items.AddRange(New DevExpress.XtraBars.BarItem() {Me.bbSave, Me.bbNew, Me.bbDelete, Me.bbClose, Me.bbShowAll, Me.bbDisplaySelected, Me.bbRefresh, Me.bbPrint})
        Me.BarManager1.MainMenu = Me.Bar2
        Me.BarManager1.MaxItemId = 8
        '
        'Bar2
        '
        Me.Bar2.BarName = "Main menu"
        Me.Bar2.DockCol = 0
        Me.Bar2.DockRow = 0
        Me.Bar2.DockStyle = DevExpress.XtraBars.BarDockStyle.Top
        Me.Bar2.FloatLocation = New System.Drawing.Point(753, 551)
        Me.Bar2.LinksPersistInfo.AddRange(New DevExpress.XtraBars.LinkPersistInfo() {New DevExpress.XtraBars.LinkPersistInfo(Me.bbSave), New DevExpress.XtraBars.LinkPersistInfo(Me.bbNew), New DevExpress.XtraBars.LinkPersistInfo(Me.bbDelete), New DevExpress.XtraBars.LinkPersistInfo(Me.bbDisplaySelected), New DevExpress.XtraBars.LinkPersistInfo(Me.bbRefresh), New DevExpress.XtraBars.LinkPersistInfo(Me.bbPrint), New DevExpress.XtraBars.LinkPersistInfo(Me.bbClose)})
        Me.Bar2.OptionsBar.AllowQuickCustomization = False
        Me.Bar2.OptionsBar.MultiLine = True
        Me.Bar2.OptionsBar.UseWholeRow = True
        Me.Bar2.Text = "Main menu"
        '
        'bbSave
        '
        Me.bbSave.Caption = "Save"
        Me.bbSave.Id = 0
        Me.bbSave.ItemShortcut = New DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.S))
        Me.bbSave.Name = "bbSave"
        Me.bbSave.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        '
        'bbNew
        '
        Me.bbNew.Caption = "New"
        Me.bbNew.Id = 1
        Me.bbNew.ItemShortcut = New DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.N))
        Me.bbNew.Name = "bbNew"
        Me.bbNew.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        '
        'bbDelete
        '
        Me.bbDelete.Caption = "Delete"
        Me.bbDelete.Id = 2
        Me.bbDelete.ItemShortcut = New DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.D))
        Me.bbDelete.Name = "bbDelete"
        Me.bbDelete.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        '
        'bbDisplaySelected
        '
        Me.bbDisplaySelected.Caption = "Display Selected"
        Me.bbDisplaySelected.Id = 5
        Me.bbDisplaySelected.Name = "bbDisplaySelected"
        '
        'bbRefresh
        '
        Me.bbRefresh.Caption = "Refresh"
        Me.bbRefresh.Id = 6
        Me.bbRefresh.Name = "bbRefresh"
        '
        'bbPrint
        '
        Me.bbPrint.Caption = "bbPrint"
        Me.bbPrint.Id = 7
        Me.bbPrint.Name = "bbPrint"
        '
        'bbClose
        '
        Me.bbClose.Caption = "Close"
        Me.bbClose.Id = 3
        Me.bbClose.ItemShortcut = New DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.W))
        Me.bbClose.Name = "bbClose"
        Me.bbClose.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        '
        'bbShowAll
        '
        Me.bbShowAll.Caption = "Show All"
        Me.bbShowAll.Id = 4
        Me.bbShowAll.Name = "bbShowAll"
        '
        'xTab1
        '
        Me.xTab1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.xTab1.Location = New System.Drawing.Point(0, 24)
        Me.xTab1.Name = "xTab1"
        Me.xTab1.SelectedTabPage = Me.XtraTabPage1
        Me.xTab1.Size = New System.Drawing.Size(842, 383)
        Me.xTab1.TabIndex = 4
        Me.xTab1.TabPages.AddRange(New DevExpress.XtraTab.XtraTabPage() {Me.XtraTabPage1, Me.XtraTabPage2})
        '
        'XtraTabPage1
        '
        Me.XtraTabPage1.Controls.Add(Me.LayoutControl1)
        Me.XtraTabPage1.Name = "XtraTabPage1"
        Me.XtraTabPage1.Size = New System.Drawing.Size(833, 352)
        Me.XtraTabPage1.Text = "New Record"
        '
        'LayoutControl1
        '
        Me.LayoutControl1.AllowCustomizationMenu = False
        Me.LayoutControl1.Controls.Add(Me.ceStockType)
        Me.LayoutControl1.Controls.Add(Me.seCurrentStockValue)
        Me.LayoutControl1.Controls.Add(Me.seCurrentStockBalance)
        Me.LayoutControl1.Controls.Add(Me.lblID)
        Me.LayoutControl1.Controls.Add(Me.seOpeningStockValue)
        Me.LayoutControl1.Controls.Add(Me.seOpeningStockBalance)
        Me.LayoutControl1.Controls.Add(Me.seReorderLevel)
        Me.LayoutControl1.Controls.Add(Me.lupSupplier)
        Me.LayoutControl1.Controls.Add(Me.seSalesPrice)
        Me.LayoutControl1.Controls.Add(Me.sePurchasePrice)
        Me.LayoutControl1.Controls.Add(Me.LupStockCategory)
        Me.LayoutControl1.Controls.Add(Me.teDescription)
        Me.LayoutControl1.Controls.Add(Me.teStockCode)
        Me.LayoutControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LayoutControl1.Location = New System.Drawing.Point(0, 0)
        Me.LayoutControl1.Name = "LayoutControl1"
        Me.LayoutControl1.OptionsFocus.EnableAutoTabOrder = False
        Me.LayoutControl1.OptionsFocus.MoveFocusDirection = DevExpress.XtraLayout.MoveFocusDirection.DownThenAcross
        Me.LayoutControl1.Root = Me.LayoutControlGroup1
        Me.LayoutControl1.Size = New System.Drawing.Size(833, 352)
        Me.LayoutControl1.TabIndex = 0
        Me.LayoutControl1.Text = "LayoutControl1"
        '
        'ceStockType
        '
        Me.ceStockType.EditValue = "STOCK ITEM"
        Me.ceStockType.EnterMoveNextControl = True
        Me.ceStockType.Location = New System.Drawing.Point(121, 116)
        Me.ceStockType.Name = "ceStockType"
        Me.ceStockType.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.ceStockType.Properties.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.ceStockType.Properties.Items.AddRange(New Object() {"STOCK ITEM", "SERVICE"})
        Me.ceStockType.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        Me.ceStockType.Size = New System.Drawing.Size(291, 20)
        Me.ceStockType.StyleController = Me.LayoutControl1
        Me.ceStockType.TabIndex = 2
        ConditionValidationRule1.ConditionOperator = DevExpress.XtraEditors.DXErrorProvider.ConditionOperator.IsNotBlank
        ConditionValidationRule1.ErrorText = "Can not be Blank"
        Me.dxvpStockItems.SetValidationRule(Me.ceStockType, ConditionValidationRule1)
        '
        'seCurrentStockValue
        '
        Me.seCurrentStockValue.EditValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.seCurrentStockValue.Location = New System.Drawing.Point(344, 305)
        Me.seCurrentStockValue.Name = "seCurrentStockValue"
        Me.seCurrentStockValue.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton})
        Me.seCurrentStockValue.Properties.DisplayFormat.FormatString = "F"
        Me.seCurrentStockValue.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.seCurrentStockValue.Properties.EditFormat.FormatString = "F"
        Me.seCurrentStockValue.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.seCurrentStockValue.Size = New System.Drawing.Size(258, 20)
        Me.seCurrentStockValue.StyleController = Me.LayoutControl1
        Me.seCurrentStockValue.TabIndex = 11
        '
        'seCurrentStockBalance
        '
        Me.seCurrentStockBalance.EditValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.seCurrentStockBalance.EnterMoveNextControl = True
        Me.seCurrentStockBalance.Location = New System.Drawing.Point(344, 274)
        Me.seCurrentStockBalance.Name = "seCurrentStockBalance"
        Me.seCurrentStockBalance.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton})
        Me.seCurrentStockBalance.Properties.DisplayFormat.FormatString = "F"
        Me.seCurrentStockBalance.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.seCurrentStockBalance.Properties.EditFormat.FormatString = "F"
        Me.seCurrentStockBalance.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.seCurrentStockBalance.Size = New System.Drawing.Size(258, 20)
        Me.seCurrentStockBalance.StyleController = Me.LayoutControl1
        Me.seCurrentStockBalance.TabIndex = 10
        '
        'lblID
        '
        Me.lblID.Location = New System.Drawing.Point(7, 30)
        Me.lblID.Name = "lblID"
        Me.lblID.Size = New System.Drawing.Size(820, 13)
        Me.lblID.StyleController = Me.LayoutControl1
        Me.lblID.TabIndex = 97
        '
        'seOpeningStockValue
        '
        Me.seOpeningStockValue.EditValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.seOpeningStockValue.EnterMoveNextControl = True
        Me.seOpeningStockValue.Location = New System.Drawing.Point(537, 209)
        Me.seOpeningStockValue.Name = "seOpeningStockValue"
        Me.seOpeningStockValue.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton})
        Me.seOpeningStockValue.Properties.DisplayFormat.FormatString = "F"
        Me.seOpeningStockValue.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.seOpeningStockValue.Properties.EditFormat.FormatString = "F"
        Me.seOpeningStockValue.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.seOpeningStockValue.Size = New System.Drawing.Size(290, 20)
        Me.seOpeningStockValue.StyleController = Me.LayoutControl1
        Me.seOpeningStockValue.TabIndex = 9
        '
        'seOpeningStockBalance
        '
        Me.seOpeningStockBalance.EditValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.seOpeningStockBalance.EnterMoveNextControl = True
        Me.seOpeningStockBalance.Location = New System.Drawing.Point(537, 178)
        Me.seOpeningStockBalance.Name = "seOpeningStockBalance"
        Me.seOpeningStockBalance.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton})
        Me.seOpeningStockBalance.Properties.DisplayFormat.FormatString = "F"
        Me.seOpeningStockBalance.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.seOpeningStockBalance.Properties.EditFormat.FormatString = "F"
        Me.seOpeningStockBalance.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.seOpeningStockBalance.Size = New System.Drawing.Size(290, 20)
        Me.seOpeningStockBalance.StyleController = Me.LayoutControl1
        Me.seOpeningStockBalance.TabIndex = 8
        '
        'seReorderLevel
        '
        Me.seReorderLevel.EditValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.seReorderLevel.EnterMoveNextControl = True
        Me.seReorderLevel.Location = New System.Drawing.Point(121, 209)
        Me.seReorderLevel.Name = "seReorderLevel"
        Me.seReorderLevel.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton})
        Me.seReorderLevel.Properties.DisplayFormat.FormatString = "F"
        Me.seReorderLevel.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.seReorderLevel.Properties.EditFormat.FormatString = "F"
        Me.seReorderLevel.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.seReorderLevel.Size = New System.Drawing.Size(291, 20)
        Me.seReorderLevel.StyleController = Me.LayoutControl1
        Me.seReorderLevel.TabIndex = 5
        '
        'lupSupplier
        '
        Me.lupSupplier.EnterMoveNextControl = True
        Me.lupSupplier.Location = New System.Drawing.Point(121, 178)
        Me.lupSupplier.Name = "lupSupplier"
        Me.lupSupplier.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.lupSupplier.Properties.Columns.AddRange(New DevExpress.XtraEditors.Controls.LookUpColumnInfo() {New DevExpress.XtraEditors.Controls.LookUpColumnInfo("SupplierNo", "Supplier No", 20, DevExpress.Utils.FormatType.None, "", True, DevExpress.Utils.HorzAlignment.[Default], DevExpress.Data.ColumnSortOrder.None), New DevExpress.XtraEditors.Controls.LookUpColumnInfo("SupplierName", "Supplier Name", 20, DevExpress.Utils.FormatType.None, "", True, DevExpress.Utils.HorzAlignment.[Default], DevExpress.Data.ColumnSortOrder.None)})
        Me.lupSupplier.Properties.NullText = ""
        Me.lupSupplier.Size = New System.Drawing.Size(291, 20)
        Me.lupSupplier.StyleController = Me.LayoutControl1
        Me.lupSupplier.TabIndex = 4
        ConditionValidationRule2.ConditionOperator = DevExpress.XtraEditors.DXErrorProvider.ConditionOperator.IsNotBlank
        ConditionValidationRule2.ErrorText = "Can not be Blank"
        Me.dxvpStockItems.SetValidationRule(Me.lupSupplier, ConditionValidationRule2)
        '
        'seSalesPrice
        '
        Me.seSalesPrice.EditValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.seSalesPrice.EnterMoveNextControl = True
        Me.seSalesPrice.Location = New System.Drawing.Point(537, 147)
        Me.seSalesPrice.Name = "seSalesPrice"
        Me.seSalesPrice.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton})
        Me.seSalesPrice.Properties.DisplayFormat.FormatString = "F"
        Me.seSalesPrice.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.seSalesPrice.Properties.EditFormat.FormatString = "F"
        Me.seSalesPrice.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.seSalesPrice.Size = New System.Drawing.Size(290, 20)
        Me.seSalesPrice.StyleController = Me.LayoutControl1
        Me.seSalesPrice.TabIndex = 7
        '
        'sePurchasePrice
        '
        Me.sePurchasePrice.EditValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.sePurchasePrice.EnterMoveNextControl = True
        Me.sePurchasePrice.Location = New System.Drawing.Point(537, 116)
        Me.sePurchasePrice.Name = "sePurchasePrice"
        Me.sePurchasePrice.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton})
        Me.sePurchasePrice.Properties.DisplayFormat.FormatString = "F"
        Me.sePurchasePrice.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.sePurchasePrice.Properties.EditFormat.FormatString = "F"
        Me.sePurchasePrice.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.sePurchasePrice.Size = New System.Drawing.Size(290, 20)
        Me.sePurchasePrice.StyleController = Me.LayoutControl1
        Me.sePurchasePrice.TabIndex = 6
        '
        'LupStockCategory
        '
        Me.LupStockCategory.EnterMoveNextControl = True
        Me.LupStockCategory.Location = New System.Drawing.Point(121, 147)
        Me.LupStockCategory.Name = "LupStockCategory"
        Me.LupStockCategory.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo), New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Glyph, "Save", -1, True, True, False, DevExpress.XtraEditors.ImageLocation.MiddleCenter, Nothing), New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Glyph, "Delete", -1, True, True, False, DevExpress.XtraEditors.ImageLocation.MiddleCenter, Nothing)})
        Me.LupStockCategory.Properties.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.LupStockCategory.Properties.Columns.AddRange(New DevExpress.XtraEditors.Controls.LookUpColumnInfo() {New DevExpress.XtraEditors.Controls.LookUpColumnInfo("Name", "Stock Category", 20, DevExpress.Utils.FormatType.None, "", True, DevExpress.Utils.HorzAlignment.[Default], DevExpress.Data.ColumnSortOrder.None)})
        Me.LupStockCategory.Properties.NullText = ""
        Me.LupStockCategory.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard
        Me.LupStockCategory.Size = New System.Drawing.Size(296, 20)
        Me.LupStockCategory.StyleController = Me.LayoutControl1
        Me.LupStockCategory.TabIndex = 3
        ConditionValidationRule3.ConditionOperator = DevExpress.XtraEditors.DXErrorProvider.ConditionOperator.IsNotBlank
        ConditionValidationRule3.ErrorText = "Can not be Blank"
        Me.dxvpStockItems.SetValidationRule(Me.LupStockCategory, ConditionValidationRule3)
        '
        'teDescription
        '
        Me.teDescription.EnterMoveNextControl = True
        Me.teDescription.Location = New System.Drawing.Point(121, 85)
        Me.teDescription.Name = "teDescription"
        Me.teDescription.Properties.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.teDescription.Size = New System.Drawing.Size(291, 20)
        Me.teDescription.StyleController = Me.LayoutControl1
        Me.teDescription.TabIndex = 1
        ConditionValidatonRule1.ConditionOperator = DevExpress.XtraEditors.DXErrorProvider.ConditionOperator.IsNotBlank
        ConditionValidatonRule1.ErrorText = "Can not be blank"
        Me.dxvpStockItems.SetValidationRule(Me.teDescription, ConditionValidatonRule1)
        '
        'teStockCode
        '
        Me.teStockCode.EnterMoveNextControl = True
        Me.teStockCode.Location = New System.Drawing.Point(121, 54)
        Me.teStockCode.Name = "teStockCode"
        Me.teStockCode.Properties.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.teStockCode.Size = New System.Drawing.Size(82, 20)
        Me.teStockCode.StyleController = Me.LayoutControl1
        Me.teStockCode.TabIndex = 0
        ConditionValidatonRule2.ConditionOperator = DevExpress.XtraEditors.DXErrorProvider.ConditionOperator.IsNotBlank
        ConditionValidatonRule2.ErrorText = "Can not be blank"
        Me.dxvpStockItems.SetValidationRule(Me.teStockCode, ConditionValidatonRule2)
        '
        'LayoutControlGroup1
        '
        Me.LayoutControlGroup1.CustomizationFormText = "LayoutControlGroup1"
        Me.LayoutControlGroup1.Items.AddRange(New DevExpress.XtraLayout.BaseLayoutItem() {Me.LayoutControlItem1, Me.LayoutControlItem2, Me.LayoutControlItem10, Me.EmptySpaceItem1, Me.EmptySpaceItem4, Me.LayoutControlItem3, Me.EmptySpaceItem3, Me.LayoutControlItem13, Me.LayoutControlItem4, Me.LayoutControlItem5, Me.LayoutControlItem6, Me.LayoutControlItem9, Me.LayoutControlItem8, Me.LayoutControlItem7, Me.EmptySpaceItem2, Me.EmptySpaceItem5, Me.EmptySpaceItem6, Me.EmptySpaceItem7, Me.EmptySpaceItem8, Me.LayoutControlGroup2, Me.EmptySpaceItem9})
        Me.LayoutControlGroup1.Location = New System.Drawing.Point(0, 0)
        Me.LayoutControlGroup1.Name = "LayoutControlGroup1"
        Me.LayoutControlGroup1.Size = New System.Drawing.Size(833, 352)
        Me.LayoutControlGroup1.Spacing = New DevExpress.XtraLayout.Utils.Padding(0, 0, 0, 0)
        Me.LayoutControlGroup1.Text = "LayoutControlGroup1"
        Me.LayoutControlGroup1.TextVisible = False
        '
        'LayoutControlItem1
        '
        Me.LayoutControlItem1.Control = Me.teStockCode
        Me.LayoutControlItem1.CustomizationFormText = "Stock Code"
        Me.LayoutControlItem1.Location = New System.Drawing.Point(0, 47)
        Me.LayoutControlItem1.Name = "LayoutControlItem1"
        Me.LayoutControlItem1.Size = New System.Drawing.Size(207, 31)
        Me.LayoutControlItem1.Text = "Stock Code"
        Me.LayoutControlItem1.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem1.TextSize = New System.Drawing.Size(109, 20)
        '
        'LayoutControlItem2
        '
        Me.LayoutControlItem2.Control = Me.teDescription
        Me.LayoutControlItem2.CustomizationFormText = "Description"
        Me.LayoutControlItem2.Location = New System.Drawing.Point(0, 78)
        Me.LayoutControlItem2.Name = "LayoutControlItem2"
        Me.LayoutControlItem2.Size = New System.Drawing.Size(416, 31)
        Me.LayoutControlItem2.Text = "Description"
        Me.LayoutControlItem2.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem2.TextSize = New System.Drawing.Size(109, 20)
        '
        'LayoutControlItem10
        '
        Me.LayoutControlItem10.Control = Me.lblID
        Me.LayoutControlItem10.CustomizationFormText = "LayoutControlItem10"
        Me.LayoutControlItem10.Location = New System.Drawing.Point(0, 23)
        Me.LayoutControlItem10.Name = "LayoutControlItem10"
        Me.LayoutControlItem10.Size = New System.Drawing.Size(831, 24)
        Me.LayoutControlItem10.Text = "LayoutControlItem10"
        Me.LayoutControlItem10.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem10.TextSize = New System.Drawing.Size(0, 0)
        Me.LayoutControlItem10.TextToControlDistance = 0
        Me.LayoutControlItem10.TextVisible = False
        Me.LayoutControlItem10.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never
        '
        'EmptySpaceItem1
        '
        Me.EmptySpaceItem1.CustomizationFormText = "EmptySpaceItem1"
        Me.EmptySpaceItem1.Location = New System.Drawing.Point(0, 0)
        Me.EmptySpaceItem1.MaxSize = New System.Drawing.Size(0, 25)
        Me.EmptySpaceItem1.MinSize = New System.Drawing.Size(10, 10)
        Me.EmptySpaceItem1.Name = "EmptySpaceItem1"
        Me.EmptySpaceItem1.Size = New System.Drawing.Size(831, 23)
        Me.EmptySpaceItem1.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom
        Me.EmptySpaceItem1.Text = "EmptySpaceItem1"
        Me.EmptySpaceItem1.TextSize = New System.Drawing.Size(0, 0)
        '
        'EmptySpaceItem4
        '
        Me.EmptySpaceItem4.CustomizationFormText = "EmptySpaceItem4"
        Me.EmptySpaceItem4.Location = New System.Drawing.Point(207, 47)
        Me.EmptySpaceItem4.Name = "EmptySpaceItem4"
        Me.EmptySpaceItem4.Size = New System.Drawing.Size(624, 31)
        Me.EmptySpaceItem4.Text = "EmptySpaceItem4"
        Me.EmptySpaceItem4.TextSize = New System.Drawing.Size(0, 0)
        '
        'LayoutControlItem3
        '
        Me.LayoutControlItem3.Control = Me.LupStockCategory
        Me.LayoutControlItem3.CustomizationFormText = "Stock Category"
        Me.LayoutControlItem3.Location = New System.Drawing.Point(0, 140)
        Me.LayoutControlItem3.Name = "LayoutControlItem3"
        Me.LayoutControlItem3.Padding = New DevExpress.XtraLayout.Utils.Padding(5, 0, 5, 5)
        Me.LayoutControlItem3.Size = New System.Drawing.Size(416, 31)
        Me.LayoutControlItem3.Text = "Stock Category"
        Me.LayoutControlItem3.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem3.TextSize = New System.Drawing.Size(109, 20)
        '
        'EmptySpaceItem3
        '
        Me.EmptySpaceItem3.CustomizationFormText = "EmptySpaceItem3"
        Me.EmptySpaceItem3.Location = New System.Drawing.Point(416, 78)
        Me.EmptySpaceItem3.Name = "EmptySpaceItem3"
        Me.EmptySpaceItem3.Size = New System.Drawing.Size(415, 31)
        Me.EmptySpaceItem3.Text = "EmptySpaceItem3"
        Me.EmptySpaceItem3.TextSize = New System.Drawing.Size(0, 0)
        '
        'LayoutControlItem13
        '
        Me.LayoutControlItem13.Control = Me.ceStockType
        Me.LayoutControlItem13.CustomizationFormText = "Type"
        Me.LayoutControlItem13.Location = New System.Drawing.Point(0, 109)
        Me.LayoutControlItem13.Name = "LayoutControlItem13"
        Me.LayoutControlItem13.Size = New System.Drawing.Size(416, 31)
        Me.LayoutControlItem13.Text = "Type"
        Me.LayoutControlItem13.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem13.TextSize = New System.Drawing.Size(109, 20)
        '
        'LayoutControlItem4
        '
        Me.LayoutControlItem4.Control = Me.sePurchasePrice
        Me.LayoutControlItem4.CustomizationFormText = "Purchasing Price"
        Me.LayoutControlItem4.Location = New System.Drawing.Point(416, 109)
        Me.LayoutControlItem4.Name = "LayoutControlItem4"
        Me.LayoutControlItem4.Size = New System.Drawing.Size(415, 31)
        Me.LayoutControlItem4.Text = "Purchasing Price"
        Me.LayoutControlItem4.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem4.TextSize = New System.Drawing.Size(109, 20)
        '
        'LayoutControlItem5
        '
        Me.LayoutControlItem5.Control = Me.seSalesPrice
        Me.LayoutControlItem5.CustomizationFormText = "Sales Price"
        Me.LayoutControlItem5.Location = New System.Drawing.Point(416, 140)
        Me.LayoutControlItem5.Name = "LayoutControlItem5"
        Me.LayoutControlItem5.Size = New System.Drawing.Size(415, 31)
        Me.LayoutControlItem5.Text = "Sales Price"
        Me.LayoutControlItem5.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem5.TextSize = New System.Drawing.Size(109, 20)
        '
        'LayoutControlItem6
        '
        Me.LayoutControlItem6.Control = Me.lupSupplier
        Me.LayoutControlItem6.CustomizationFormText = "Supplier"
        Me.LayoutControlItem6.Location = New System.Drawing.Point(0, 171)
        Me.LayoutControlItem6.Name = "LayoutControlItem6"
        Me.LayoutControlItem6.Size = New System.Drawing.Size(416, 31)
        Me.LayoutControlItem6.Text = "Supplier"
        Me.LayoutControlItem6.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem6.TextSize = New System.Drawing.Size(109, 20)
        '
        'LayoutControlItem9
        '
        Me.LayoutControlItem9.Control = Me.seOpeningStockValue
        Me.LayoutControlItem9.CustomizationFormText = "Stock Value"
        Me.LayoutControlItem9.Location = New System.Drawing.Point(416, 202)
        Me.LayoutControlItem9.Name = "LayoutControlItem9"
        Me.LayoutControlItem9.Size = New System.Drawing.Size(415, 31)
        Me.LayoutControlItem9.Text = "Opening Stock Value"
        Me.LayoutControlItem9.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem9.TextSize = New System.Drawing.Size(109, 20)
        '
        'LayoutControlItem8
        '
        Me.LayoutControlItem8.Control = Me.seOpeningStockBalance
        Me.LayoutControlItem8.CustomizationFormText = "Stock Balance"
        Me.LayoutControlItem8.Location = New System.Drawing.Point(416, 171)
        Me.LayoutControlItem8.Name = "LayoutControlItem8"
        Me.LayoutControlItem8.Size = New System.Drawing.Size(415, 31)
        Me.LayoutControlItem8.Text = "Opening Stock Balance"
        Me.LayoutControlItem8.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem8.TextSize = New System.Drawing.Size(109, 20)
        '
        'LayoutControlItem7
        '
        Me.LayoutControlItem7.Control = Me.seReorderLevel
        Me.LayoutControlItem7.CustomizationFormText = "Reorder Level"
        Me.LayoutControlItem7.Location = New System.Drawing.Point(0, 202)
        Me.LayoutControlItem7.Name = "LayoutControlItem7"
        Me.LayoutControlItem7.Size = New System.Drawing.Size(416, 31)
        Me.LayoutControlItem7.Text = "Reorder Level"
        Me.LayoutControlItem7.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem7.TextSize = New System.Drawing.Size(109, 20)
        '
        'EmptySpaceItem2
        '
        Me.EmptySpaceItem2.CustomizationFormText = "EmptySpaceItem2"
        Me.EmptySpaceItem2.Location = New System.Drawing.Point(0, 264)
        Me.EmptySpaceItem2.Name = "EmptySpaceItem2"
        Me.EmptySpaceItem2.Size = New System.Drawing.Size(220, 25)
        Me.EmptySpaceItem2.Text = "EmptySpaceItem2"
        Me.EmptySpaceItem2.TextSize = New System.Drawing.Size(0, 0)
        '
        'EmptySpaceItem5
        '
        Me.EmptySpaceItem5.CustomizationFormText = "EmptySpaceItem5"
        Me.EmptySpaceItem5.Location = New System.Drawing.Point(0, 289)
        Me.EmptySpaceItem5.Name = "EmptySpaceItem5"
        Me.EmptySpaceItem5.Size = New System.Drawing.Size(220, 43)
        Me.EmptySpaceItem5.Text = "EmptySpaceItem5"
        Me.EmptySpaceItem5.TextSize = New System.Drawing.Size(0, 0)
        '
        'EmptySpaceItem6
        '
        Me.EmptySpaceItem6.CustomizationFormText = "EmptySpaceItem6"
        Me.EmptySpaceItem6.Location = New System.Drawing.Point(609, 264)
        Me.EmptySpaceItem6.Name = "EmptySpaceItem6"
        Me.EmptySpaceItem6.Size = New System.Drawing.Size(222, 25)
        Me.EmptySpaceItem6.Text = "EmptySpaceItem6"
        Me.EmptySpaceItem6.TextSize = New System.Drawing.Size(0, 0)
        '
        'EmptySpaceItem7
        '
        Me.EmptySpaceItem7.CustomizationFormText = "EmptySpaceItem7"
        Me.EmptySpaceItem7.Location = New System.Drawing.Point(609, 289)
        Me.EmptySpaceItem7.Name = "EmptySpaceItem7"
        Me.EmptySpaceItem7.Size = New System.Drawing.Size(222, 43)
        Me.EmptySpaceItem7.Text = "EmptySpaceItem7"
        Me.EmptySpaceItem7.TextSize = New System.Drawing.Size(0, 0)
        '
        'EmptySpaceItem8
        '
        Me.EmptySpaceItem8.CustomizationFormText = "EmptySpaceItem8"
        Me.EmptySpaceItem8.Location = New System.Drawing.Point(0, 233)
        Me.EmptySpaceItem8.MaxSize = New System.Drawing.Size(0, 31)
        Me.EmptySpaceItem8.MinSize = New System.Drawing.Size(10, 31)
        Me.EmptySpaceItem8.Name = "EmptySpaceItem8"
        Me.EmptySpaceItem8.Size = New System.Drawing.Size(831, 31)
        Me.EmptySpaceItem8.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom
        Me.EmptySpaceItem8.Text = "EmptySpaceItem8"
        Me.EmptySpaceItem8.TextSize = New System.Drawing.Size(0, 0)
        '
        'LayoutControlGroup2
        '
        Me.LayoutControlGroup2.CustomizationFormText = "LayoutControlGroup2"
        Me.LayoutControlGroup2.Items.AddRange(New DevExpress.XtraLayout.BaseLayoutItem() {Me.LayoutControlItem11, Me.LayoutControlItem12})
        Me.LayoutControlGroup2.Location = New System.Drawing.Point(220, 264)
        Me.LayoutControlGroup2.Name = "LayoutControlGroup2"
        Me.LayoutControlGroup2.Size = New System.Drawing.Size(389, 68)
        Me.LayoutControlGroup2.Text = "LayoutControlGroup2"
        Me.LayoutControlGroup2.TextVisible = False
        '
        'LayoutControlItem11
        '
        Me.LayoutControlItem11.Control = Me.seCurrentStockBalance
        Me.LayoutControlItem11.CustomizationFormText = "Current Stock Balance"
        Me.LayoutControlItem11.Location = New System.Drawing.Point(0, 0)
        Me.LayoutControlItem11.Name = "LayoutControlItem11"
        Me.LayoutControlItem11.Size = New System.Drawing.Size(383, 31)
        Me.LayoutControlItem11.Text = "Current Stock Balance"
        Me.LayoutControlItem11.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem11.TextSize = New System.Drawing.Size(109, 20)
        '
        'LayoutControlItem12
        '
        Me.LayoutControlItem12.Control = Me.seCurrentStockValue
        Me.LayoutControlItem12.CustomizationFormText = "Current Stock Value"
        Me.LayoutControlItem12.Location = New System.Drawing.Point(0, 31)
        Me.LayoutControlItem12.Name = "LayoutControlItem12"
        Me.LayoutControlItem12.Size = New System.Drawing.Size(383, 31)
        Me.LayoutControlItem12.Text = "Current Stock Value"
        Me.LayoutControlItem12.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem12.TextSize = New System.Drawing.Size(109, 20)
        '
        'EmptySpaceItem9
        '
        Me.EmptySpaceItem9.CustomizationFormText = "EmptySpaceItem9"
        Me.EmptySpaceItem9.Location = New System.Drawing.Point(0, 332)
        Me.EmptySpaceItem9.Name = "EmptySpaceItem9"
        Me.EmptySpaceItem9.Size = New System.Drawing.Size(831, 18)
        Me.EmptySpaceItem9.Text = "EmptySpaceItem9"
        Me.EmptySpaceItem9.TextSize = New System.Drawing.Size(0, 0)
        '
        'XtraTabPage2
        '
        Me.XtraTabPage2.Controls.Add(Me.gcStockItems)
        Me.XtraTabPage2.Name = "XtraTabPage2"
        Me.XtraTabPage2.Size = New System.Drawing.Size(833, 352)
        Me.XtraTabPage2.Text = "History Data"
        '
        'gcStockItems
        '
        Me.gcStockItems.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gcStockItems.Location = New System.Drawing.Point(0, 0)
        Me.gcStockItems.MainView = Me.gvStockItems
        Me.gcStockItems.Name = "gcStockItems"
        Me.gcStockItems.Size = New System.Drawing.Size(833, 352)
        Me.gcStockItems.TabIndex = 0
        Me.gcStockItems.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.gvStockItems})
        '
        'gvStockItems
        '
        Me.gvStockItems.Columns.AddRange(New DevExpress.XtraGrid.Columns.GridColumn() {Me.GridColumn1, Me.GridColumn10, Me.GridColumn2, Me.GridColumn17, Me.GridColumn3, Me.GridColumn4, Me.GridColumn5, Me.GridColumn6, Me.GridColumn7, Me.GridColumn8, Me.GridColumn9, Me.GridColumn11, Me.GridColumn12, Me.GridColumn13, Me.GridColumn14, Me.GridColumn15, Me.GridColumn16})
        Me.gvStockItems.GridControl = Me.gcStockItems
        Me.gvStockItems.Name = "gvStockItems"
        Me.gvStockItems.OptionsBehavior.Editable = False
        Me.gvStockItems.OptionsView.ColumnAutoWidth = False
        Me.gvStockItems.OptionsView.ShowAutoFilterRow = True
        Me.gvStockItems.OptionsView.ShowFooter = True
        '
        'GridColumn1
        '
        Me.GridColumn1.Caption = "StockID"
        Me.GridColumn1.FieldName = "StockID"
        Me.GridColumn1.Name = "GridColumn1"
        '
        'GridColumn10
        '
        Me.GridColumn10.Caption = "Stock Code"
        Me.GridColumn10.FieldName = "StockCode"
        Me.GridColumn10.Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left
        Me.GridColumn10.Name = "GridColumn10"
        Me.GridColumn10.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Count
        Me.GridColumn10.Visible = True
        Me.GridColumn10.VisibleIndex = 0
        Me.GridColumn10.Width = 120
        '
        'GridColumn2
        '
        Me.GridColumn2.Caption = "Description"
        Me.GridColumn2.FieldName = "Description"
        Me.GridColumn2.Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left
        Me.GridColumn2.Name = "GridColumn2"
        Me.GridColumn2.Visible = True
        Me.GridColumn2.VisibleIndex = 1
        Me.GridColumn2.Width = 150
        '
        'GridColumn17
        '
        Me.GridColumn17.Caption = "Stock Type"
        Me.GridColumn17.FieldName = "StockType"
        Me.GridColumn17.Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left
        Me.GridColumn17.Name = "GridColumn17"
        Me.GridColumn17.Visible = True
        Me.GridColumn17.VisibleIndex = 2
        Me.GridColumn17.Width = 120
        '
        'GridColumn3
        '
        Me.GridColumn3.Caption = "Stock Category"
        Me.GridColumn3.FieldName = "CatagoryName"
        Me.GridColumn3.Name = "GridColumn3"
        Me.GridColumn3.Visible = True
        Me.GridColumn3.VisibleIndex = 5
        Me.GridColumn3.Width = 120
        '
        'GridColumn4
        '
        Me.GridColumn4.Caption = "Purchasing Price"
        Me.GridColumn4.DisplayFormat.FormatString = "F"
        Me.GridColumn4.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.GridColumn4.FieldName = "PurchasingPrice"
        Me.GridColumn4.Name = "GridColumn4"
        Me.GridColumn4.Visible = True
        Me.GridColumn4.VisibleIndex = 6
        Me.GridColumn4.Width = 120
        '
        'GridColumn5
        '
        Me.GridColumn5.Caption = "Sales Price"
        Me.GridColumn5.DisplayFormat.FormatString = "F"
        Me.GridColumn5.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.GridColumn5.FieldName = "SalesPrice"
        Me.GridColumn5.Name = "GridColumn5"
        Me.GridColumn5.Visible = True
        Me.GridColumn5.VisibleIndex = 7
        Me.GridColumn5.Width = 120
        '
        'GridColumn6
        '
        Me.GridColumn6.Caption = "Supplier"
        Me.GridColumn6.FieldName = "SupplierName"
        Me.GridColumn6.Name = "GridColumn6"
        Me.GridColumn6.Visible = True
        Me.GridColumn6.VisibleIndex = 8
        Me.GridColumn6.Width = 120
        '
        'GridColumn7
        '
        Me.GridColumn7.Caption = "Reorder Level"
        Me.GridColumn7.DisplayFormat.FormatString = "F"
        Me.GridColumn7.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.GridColumn7.FieldName = "ReorderLevel"
        Me.GridColumn7.Name = "GridColumn7"
        Me.GridColumn7.Visible = True
        Me.GridColumn7.VisibleIndex = 9
        Me.GridColumn7.Width = 120
        '
        'GridColumn8
        '
        Me.GridColumn8.Caption = "Opening Stock Balance"
        Me.GridColumn8.DisplayFormat.FormatString = "F"
        Me.GridColumn8.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.GridColumn8.FieldName = "OpeningStockBalance"
        Me.GridColumn8.Name = "GridColumn8"
        Me.GridColumn8.Width = 130
        '
        'GridColumn9
        '
        Me.GridColumn9.Caption = "Opening Stock Value"
        Me.GridColumn9.DisplayFormat.FormatString = "F"
        Me.GridColumn9.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.GridColumn9.FieldName = "OpeningStockValue"
        Me.GridColumn9.Name = "GridColumn9"
        Me.GridColumn9.Width = 120
        '
        'GridColumn11
        '
        Me.GridColumn11.Caption = "Current Stock Balance"
        Me.GridColumn11.FieldName = "CurrentStockBalance"
        Me.GridColumn11.Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left
        Me.GridColumn11.Name = "GridColumn11"
        Me.GridColumn11.Visible = True
        Me.GridColumn11.VisibleIndex = 3
        Me.GridColumn11.Width = 120
        '
        'GridColumn12
        '
        Me.GridColumn12.Caption = "Current Stock Value"
        Me.GridColumn12.FieldName = "CurrentStockValue"
        Me.GridColumn12.Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left
        Me.GridColumn12.Name = "GridColumn12"
        Me.GridColumn12.Visible = True
        Me.GridColumn12.VisibleIndex = 4
        Me.GridColumn12.Width = 120
        '
        'GridColumn13
        '
        Me.GridColumn13.Caption = "Created By"
        Me.GridColumn13.FieldName = "CreatedBy"
        Me.GridColumn13.Name = "GridColumn13"
        Me.GridColumn13.Visible = True
        Me.GridColumn13.VisibleIndex = 10
        Me.GridColumn13.Width = 120
        '
        'GridColumn14
        '
        Me.GridColumn14.Caption = "Created Date"
        Me.GridColumn14.DisplayFormat.FormatString = "dd-MMM-yy h:mm tt"
        Me.GridColumn14.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime
        Me.GridColumn14.FieldName = "CreatedDate"
        Me.GridColumn14.FilterMode = DevExpress.XtraGrid.ColumnFilterMode.DisplayText
        Me.GridColumn14.Name = "GridColumn14"
        Me.GridColumn14.Visible = True
        Me.GridColumn14.VisibleIndex = 11
        Me.GridColumn14.Width = 120
        '
        'GridColumn15
        '
        Me.GridColumn15.Caption = "Updated By"
        Me.GridColumn15.FieldName = "UpdatedBy"
        Me.GridColumn15.Name = "GridColumn15"
        Me.GridColumn15.Visible = True
        Me.GridColumn15.VisibleIndex = 12
        Me.GridColumn15.Width = 120
        '
        'GridColumn16
        '
        Me.GridColumn16.Caption = "Updated Date"
        Me.GridColumn16.DisplayFormat.FormatString = "dd-MMM-yy h:mm tt"
        Me.GridColumn16.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime
        Me.GridColumn16.FieldName = "UpdatedDate"
        Me.GridColumn16.FilterMode = DevExpress.XtraGrid.ColumnFilterMode.DisplayText
        Me.GridColumn16.Name = "GridColumn16"
        Me.GridColumn16.Visible = True
        Me.GridColumn16.VisibleIndex = 13
        Me.GridColumn16.Width = 120
        '
        'frmStock
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(842, 407)
        Me.Controls.Add(Me.xTab1)
        Me.Controls.Add(Me.barDockControlLeft)
        Me.Controls.Add(Me.barDockControlRight)
        Me.Controls.Add(Me.barDockControlBottom)
        Me.Controls.Add(Me.barDockControlTop)
        Me.KeyPreview = True
        Me.Name = "frmStock"
        Me.Text = "Stock Items"
        CType(Me.BarManager1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.xTab1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.xTab1.ResumeLayout(False)
        Me.XtraTabPage1.ResumeLayout(False)
        CType(Me.LayoutControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.LayoutControl1.ResumeLayout(False)
        CType(Me.ceStockType.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.seCurrentStockValue.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.seCurrentStockBalance.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.seOpeningStockValue.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.seOpeningStockBalance.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.seReorderLevel.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lupSupplier.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.seSalesPrice.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.sePurchasePrice.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LupStockCategory.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.teDescription.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.teStockCode.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlGroup1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem10, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.EmptySpaceItem1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.EmptySpaceItem4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.EmptySpaceItem3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem13, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem9, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem8, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.EmptySpaceItem2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.EmptySpaceItem5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.EmptySpaceItem6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.EmptySpaceItem7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.EmptySpaceItem8, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlGroup2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem11, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem12, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.EmptySpaceItem9, System.ComponentModel.ISupportInitialize).EndInit()
        Me.XtraTabPage2.ResumeLayout(False)
        CType(Me.gcStockItems, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gvStockItems, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dxvpStockItems, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents BarManager1 As DevExpress.XtraBars.BarManager
    Friend WithEvents Bar2 As DevExpress.XtraBars.Bar
    Friend WithEvents bbSave As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents bbNew As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents bbDelete As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents bbClose As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents barDockControlTop As DevExpress.XtraBars.BarDockControl
    Friend WithEvents barDockControlBottom As DevExpress.XtraBars.BarDockControl
    Friend WithEvents barDockControlLeft As DevExpress.XtraBars.BarDockControl
    Friend WithEvents barDockControlRight As DevExpress.XtraBars.BarDockControl
    Friend WithEvents bbShowAll As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents bbDisplaySelected As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents bbRefresh As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents xTab1 As DevExpress.XtraTab.XtraTabControl
    Friend WithEvents XtraTabPage1 As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents XtraTabPage2 As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents gcStockItems As DevExpress.XtraGrid.GridControl
    Friend WithEvents gvStockItems As DevExpress.XtraGrid.Views.Grid.GridView
    Friend WithEvents GridColumn1 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn2 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn3 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn4 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn5 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn6 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn7 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn8 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn9 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents dxvpStockItems As DevExpress.XtraEditors.DXErrorProvider.DXValidationProvider
    Friend WithEvents GridColumn10 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents LayoutControl1 As DevExpress.XtraLayout.LayoutControl
    Friend WithEvents LayoutControlGroup1 As DevExpress.XtraLayout.LayoutControlGroup
    Friend WithEvents teStockCode As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LayoutControlItem1 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents teDescription As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LayoutControlItem2 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents LupStockCategory As DevExpress.XtraEditors.LookUpEdit
    Friend WithEvents LayoutControlItem3 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents sePurchasePrice As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents LayoutControlItem4 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents seSalesPrice As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents LayoutControlItem5 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents lupSupplier As DevExpress.XtraEditors.LookUpEdit
    Friend WithEvents LayoutControlItem6 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents seReorderLevel As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents LayoutControlItem7 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents seOpeningStockBalance As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents LayoutControlItem8 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents seOpeningStockValue As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents LayoutControlItem9 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents lblID As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LayoutControlItem10 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents EmptySpaceItem1 As DevExpress.XtraLayout.EmptySpaceItem
    Friend WithEvents EmptySpaceItem3 As DevExpress.XtraLayout.EmptySpaceItem
    Friend WithEvents bbPrint As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents seCurrentStockValue As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents seCurrentStockBalance As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents LayoutControlItem11 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents LayoutControlItem12 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents GridColumn11 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn12 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn13 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn14 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn15 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn16 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents ceStockType As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents LayoutControlItem13 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents GridColumn17 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents EmptySpaceItem4 As DevExpress.XtraLayout.EmptySpaceItem
    Friend WithEvents EmptySpaceItem2 As DevExpress.XtraLayout.EmptySpaceItem
    Friend WithEvents EmptySpaceItem5 As DevExpress.XtraLayout.EmptySpaceItem
    Friend WithEvents EmptySpaceItem6 As DevExpress.XtraLayout.EmptySpaceItem
    Friend WithEvents EmptySpaceItem7 As DevExpress.XtraLayout.EmptySpaceItem
    Friend WithEvents EmptySpaceItem8 As DevExpress.XtraLayout.EmptySpaceItem
    Friend WithEvents LayoutControlGroup2 As DevExpress.XtraLayout.LayoutControlGroup
    Friend WithEvents EmptySpaceItem9 As DevExpress.XtraLayout.EmptySpaceItem
End Class
