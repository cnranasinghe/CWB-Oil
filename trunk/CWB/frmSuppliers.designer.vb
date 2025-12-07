<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSuppliers
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
        Me.xTab1 = New DevExpress.XtraTab.XtraTabControl
        Me.XtraTabPage1 = New DevExpress.XtraTab.XtraTabPage
        Me.LayoutControl1 = New DevExpress.XtraLayout.LayoutControl
        Me.seSupplierNo = New DevExpress.XtraEditors.SpinEdit
        Me.cbeSalutation = New DevExpress.XtraEditors.ComboBoxEdit
        Me.ceBank = New DevExpress.XtraEditors.ComboBoxEdit
        Me.teAddressLine3 = New DevExpress.XtraEditors.TextEdit
        Me.teAddressLine2 = New DevExpress.XtraEditors.TextEdit
        Me.lblCustomerID = New DevExpress.XtraEditors.LabelControl
        Me.teAccountNo = New DevExpress.XtraEditors.TextEdit
        Me.teTelephone = New DevExpress.XtraEditors.TextEdit
        Me.teEstate = New DevExpress.XtraEditors.TextEdit
        Me.teAddressLine1 = New DevExpress.XtraEditors.TextEdit
        Me.teSupplierName = New DevExpress.XtraEditors.TextEdit
        Me.LayoutControlGroup1 = New DevExpress.XtraLayout.LayoutControlGroup
        Me.LayoutControlItem3 = New DevExpress.XtraLayout.LayoutControlItem
        Me.LayoutControlItem8 = New DevExpress.XtraLayout.LayoutControlItem
        Me.LayoutControlItem6 = New DevExpress.XtraLayout.LayoutControlItem
        Me.LayoutControlItem9 = New DevExpress.XtraLayout.LayoutControlItem
        Me.LayoutControlItem10 = New DevExpress.XtraLayout.LayoutControlItem
        Me.EmptySpaceItem2 = New DevExpress.XtraLayout.EmptySpaceItem
        Me.LayoutControlItem7 = New DevExpress.XtraLayout.LayoutControlItem
        Me.EmptySpaceItem3 = New DevExpress.XtraLayout.EmptySpaceItem
        Me.LayoutControlGroup2 = New DevExpress.XtraLayout.LayoutControlGroup
        Me.LayoutControlItem2 = New DevExpress.XtraLayout.LayoutControlItem
        Me.LayoutControlItem11 = New DevExpress.XtraLayout.LayoutControlItem
        Me.LayoutControlItem12 = New DevExpress.XtraLayout.LayoutControlItem
        Me.EmptySpaceItem1 = New DevExpress.XtraLayout.EmptySpaceItem
        Me.EmptySpaceItem4 = New DevExpress.XtraLayout.EmptySpaceItem
        Me.LayoutControlItem5 = New DevExpress.XtraLayout.LayoutControlItem
        Me.LayoutControlItem4 = New DevExpress.XtraLayout.LayoutControlItem
        Me.EmptySpaceItem5 = New DevExpress.XtraLayout.EmptySpaceItem
        Me.EmptySpaceItem6 = New DevExpress.XtraLayout.EmptySpaceItem
        Me.EmptySpaceItem7 = New DevExpress.XtraLayout.EmptySpaceItem
        Me.XtraTabPage2 = New DevExpress.XtraTab.XtraTabPage
        Me.gcSuppliers = New DevExpress.XtraGrid.GridControl
        Me.gvSuppliers = New DevExpress.XtraGrid.Views.Grid.GridView
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
        Me.GridColumn11 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn12 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridView2 = New DevExpress.XtraGrid.Views.Grid.GridView
        Me.dxvpSuppliers = New DevExpress.XtraEditors.DXErrorProvider.DXValidationProvider(Me.components)
        CType(Me.BarManager1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.xTab1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.xTab1.SuspendLayout()
        Me.XtraTabPage1.SuspendLayout()
        CType(Me.LayoutControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.LayoutControl1.SuspendLayout()
        CType(Me.seSupplierNo.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cbeSalutation.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ceBank.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.teAddressLine3.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.teAddressLine2.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.teAccountNo.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.teTelephone.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.teEstate.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.teAddressLine1.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.teSupplierName.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlGroup1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem8, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem9, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem10, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.EmptySpaceItem2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.EmptySpaceItem3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlGroup2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem11, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem12, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.EmptySpaceItem1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.EmptySpaceItem4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.EmptySpaceItem5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.EmptySpaceItem6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.EmptySpaceItem7, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.XtraTabPage2.SuspendLayout()
        CType(Me.gcSuppliers, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gvSuppliers, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dxvpSuppliers, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.BarManager1.Items.AddRange(New DevExpress.XtraBars.BarItem() {Me.bbSave, Me.bbNew, Me.bbDelete, Me.bbDisplaySelected, Me.bbRefresh, Me.bbClose, Me.bbPrint})
        Me.BarManager1.MainMenu = Me.Bar2
        Me.BarManager1.MaxItemId = 7
        '
        'Bar2
        '
        Me.Bar2.BarName = "Main menu"
        Me.Bar2.DockCol = 0
        Me.Bar2.DockRow = 0
        Me.Bar2.DockStyle = DevExpress.XtraBars.BarDockStyle.Top
        Me.Bar2.LinksPersistInfo.AddRange(New DevExpress.XtraBars.LinkPersistInfo() {New DevExpress.XtraBars.LinkPersistInfo(Me.bbSave), New DevExpress.XtraBars.LinkPersistInfo(Me.bbNew), New DevExpress.XtraBars.LinkPersistInfo(Me.bbDelete), New DevExpress.XtraBars.LinkPersistInfo(Me.bbDisplaySelected), New DevExpress.XtraBars.LinkPersistInfo(Me.bbRefresh), New DevExpress.XtraBars.LinkPersistInfo(Me.bbPrint), New DevExpress.XtraBars.LinkPersistInfo(Me.bbClose)})
        Me.Bar2.OptionsBar.AllowQuickCustomization = False
        Me.Bar2.OptionsBar.MultiLine = True
        Me.Bar2.OptionsBar.UseWholeRow = True
        Me.Bar2.Text = "Main Menu"
        '
        'bbSave
        '
        Me.bbSave.Caption = "BarButtonItem1"
        Me.bbSave.Id = 0
        Me.bbSave.Name = "bbSave"
        Me.bbSave.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        '
        'bbNew
        '
        Me.bbNew.Caption = "BarButtonItem2"
        Me.bbNew.Id = 1
        Me.bbNew.Name = "bbNew"
        '
        'bbDelete
        '
        Me.bbDelete.Caption = "BarButtonItem3"
        Me.bbDelete.Id = 2
        Me.bbDelete.Name = "bbDelete"
        '
        'bbDisplaySelected
        '
        Me.bbDisplaySelected.Caption = "BarButtonItem4"
        Me.bbDisplaySelected.Id = 3
        Me.bbDisplaySelected.Name = "bbDisplaySelected"
        '
        'bbRefresh
        '
        Me.bbRefresh.Caption = "BarButtonItem5"
        Me.bbRefresh.Id = 4
        Me.bbRefresh.Name = "bbRefresh"
        '
        'bbPrint
        '
        Me.bbPrint.Caption = "BarButtonItem1"
        Me.bbPrint.Id = 6
        Me.bbPrint.Name = "bbPrint"
        '
        'bbClose
        '
        Me.bbClose.Caption = "BarButtonItem1"
        Me.bbClose.Id = 5
        Me.bbClose.Name = "bbClose"
        '
        'xTab1
        '
        Me.xTab1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.xTab1.Location = New System.Drawing.Point(0, 46)
        Me.xTab1.Name = "xTab1"
        Me.xTab1.SelectedTabPage = Me.XtraTabPage1
        Me.xTab1.Size = New System.Drawing.Size(642, 345)
        Me.xTab1.TabIndex = 4
        Me.xTab1.TabPages.AddRange(New DevExpress.XtraTab.XtraTabPage() {Me.XtraTabPage1, Me.XtraTabPage2})
        '
        'XtraTabPage1
        '
        Me.XtraTabPage1.Controls.Add(Me.LayoutControl1)
        Me.XtraTabPage1.Name = "XtraTabPage1"
        Me.XtraTabPage1.Size = New System.Drawing.Size(633, 315)
        Me.XtraTabPage1.Text = "New Record"
        '
        'LayoutControl1
        '
        Me.LayoutControl1.Controls.Add(Me.seSupplierNo)
        Me.LayoutControl1.Controls.Add(Me.cbeSalutation)
        Me.LayoutControl1.Controls.Add(Me.ceBank)
        Me.LayoutControl1.Controls.Add(Me.teAddressLine3)
        Me.LayoutControl1.Controls.Add(Me.teAddressLine2)
        Me.LayoutControl1.Controls.Add(Me.lblCustomerID)
        Me.LayoutControl1.Controls.Add(Me.teAccountNo)
        Me.LayoutControl1.Controls.Add(Me.teTelephone)
        Me.LayoutControl1.Controls.Add(Me.teEstate)
        Me.LayoutControl1.Controls.Add(Me.teAddressLine1)
        Me.LayoutControl1.Controls.Add(Me.teSupplierName)
        Me.LayoutControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LayoutControl1.Location = New System.Drawing.Point(0, 0)
        Me.LayoutControl1.Name = "LayoutControl1"
        Me.LayoutControl1.Root = Me.LayoutControlGroup1
        Me.LayoutControl1.Size = New System.Drawing.Size(633, 315)
        Me.LayoutControl1.TabIndex = 0
        Me.LayoutControl1.Text = "LayoutControl1"
        '
        'seSupplierNo
        '
        Me.seSupplierNo.EditValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.seSupplierNo.EnterMoveNextControl = True
        Me.seSupplierNo.Location = New System.Drawing.Point(82, 62)
        Me.seSupplierNo.Name = "seSupplierNo"
        Me.seSupplierNo.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton})
        Me.seSupplierNo.Properties.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.seSupplierNo.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.seSupplierNo.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.seSupplierNo.Properties.Mask.EditMask = "n0"
        Me.seSupplierNo.Properties.MaxLength = 9
        Me.seSupplierNo.Size = New System.Drawing.Size(71, 20)
        Me.seSupplierNo.StyleController = Me.LayoutControl1
        Me.seSupplierNo.TabIndex = 0
        ConditionValidationRule1.ConditionOperator = DevExpress.XtraEditors.DXErrorProvider.ConditionOperator.Greater
        ConditionValidationRule1.ErrorText = "Can not be Blank"
        ConditionValidationRule1.Value1 = 0
        Me.dxvpSuppliers.SetValidationRule(Me.seSupplierNo, ConditionValidationRule1)
        '
        'cbeSalutation
        '
        Me.cbeSalutation.EnterMoveNextControl = True
        Me.cbeSalutation.Location = New System.Drawing.Point(82, 93)
        Me.cbeSalutation.Name = "cbeSalutation"
        Me.cbeSalutation.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cbeSalutation.Properties.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.cbeSalutation.Properties.Items.AddRange(New Object() {"MR.", "MRS.", "MS.", "VEN."})
        Me.cbeSalutation.Properties.MaxLength = 5
        Me.cbeSalutation.Size = New System.Drawing.Size(50, 20)
        Me.cbeSalutation.StyleController = Me.LayoutControl1
        Me.cbeSalutation.TabIndex = 15
        '
        'ceBank
        '
        Me.ceBank.EnterMoveNextControl = True
        Me.ceBank.Location = New System.Drawing.Point(397, 248)
        Me.ceBank.Name = "ceBank"
        Me.ceBank.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.ceBank.Properties.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.ceBank.Properties.Items.AddRange(New Object() {"BOC", "CITY", "COMMERCIAL", "CORPORATIVE", "DFCC", "HDFC", "HNB", "HSBC", "NDB", "NSB", "PAN ASIA", "PEOPLE'S", "SAMURDI", "SANASA", "SEYLAN", "STANDARD CHARTED", "STATE MORTGAE"})
        Me.ceBank.Properties.MaxLength = 50
        Me.ceBank.Properties.Sorted = True
        Me.ceBank.Size = New System.Drawing.Size(230, 20)
        Me.ceBank.StyleController = Me.LayoutControl1
        Me.ceBank.TabIndex = 14
        '
        'teAddressLine3
        '
        Me.teAddressLine3.EnterMoveNextControl = True
        Me.teAddressLine3.Location = New System.Drawing.Point(82, 186)
        Me.teAddressLine3.Name = "teAddressLine3"
        Me.teAddressLine3.Properties.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.teAddressLine3.Properties.MaxLength = 100
        Me.teAddressLine3.Size = New System.Drawing.Size(229, 20)
        Me.teAddressLine3.StyleController = Me.LayoutControl1
        Me.teAddressLine3.TabIndex = 13
        '
        'teAddressLine2
        '
        Me.teAddressLine2.EnterMoveNextControl = True
        Me.teAddressLine2.Location = New System.Drawing.Point(82, 155)
        Me.teAddressLine2.Name = "teAddressLine2"
        Me.teAddressLine2.Properties.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.teAddressLine2.Properties.MaxLength = 100
        Me.teAddressLine2.Size = New System.Drawing.Size(229, 20)
        Me.teAddressLine2.StyleController = Me.LayoutControl1
        Me.teAddressLine2.TabIndex = 12
        '
        'lblCustomerID
        '
        Me.lblCustomerID.Location = New System.Drawing.Point(7, 38)
        Me.lblCustomerID.Name = "lblCustomerID"
        Me.lblCustomerID.Size = New System.Drawing.Size(620, 13)
        Me.lblCustomerID.StyleController = Me.LayoutControl1
        Me.lblCustomerID.TabIndex = 11
        '
        'teAccountNo
        '
        Me.teAccountNo.EnterMoveNextControl = True
        Me.teAccountNo.Location = New System.Drawing.Point(82, 248)
        Me.teAccountNo.Name = "teAccountNo"
        Me.teAccountNo.Properties.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.teAccountNo.Properties.MaxLength = 50
        Me.teAccountNo.Size = New System.Drawing.Size(229, 20)
        Me.teAccountNo.StyleController = Me.LayoutControl1
        Me.teAccountNo.TabIndex = 9
        '
        'teTelephone
        '
        Me.teTelephone.EnterMoveNextControl = True
        Me.teTelephone.Location = New System.Drawing.Point(82, 217)
        Me.teTelephone.Name = "teTelephone"
        Me.teTelephone.Properties.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.teTelephone.Properties.MaxLength = 50
        Me.teTelephone.Size = New System.Drawing.Size(229, 20)
        Me.teTelephone.StyleController = Me.LayoutControl1
        Me.teTelephone.TabIndex = 8
        '
        'teEstate
        '
        Me.teEstate.EnterMoveNextControl = True
        Me.teEstate.Location = New System.Drawing.Point(397, 217)
        Me.teEstate.Name = "teEstate"
        Me.teEstate.Properties.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.teEstate.Properties.MaxLength = 50
        Me.teEstate.Size = New System.Drawing.Size(230, 20)
        Me.teEstate.StyleController = Me.LayoutControl1
        Me.teEstate.TabIndex = 7
        '
        'teAddressLine1
        '
        Me.teAddressLine1.EnterMoveNextControl = True
        Me.teAddressLine1.Location = New System.Drawing.Point(82, 124)
        Me.teAddressLine1.Name = "teAddressLine1"
        Me.teAddressLine1.Properties.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.teAddressLine1.Properties.MaxLength = 100
        Me.teAddressLine1.Size = New System.Drawing.Size(229, 20)
        Me.teAddressLine1.StyleController = Me.LayoutControl1
        Me.teAddressLine1.TabIndex = 6
        '
        'teSupplierName
        '
        Me.teSupplierName.EnterMoveNextControl = True
        Me.teSupplierName.Location = New System.Drawing.Point(218, 93)
        Me.teSupplierName.Name = "teSupplierName"
        Me.teSupplierName.Properties.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.teSupplierName.Properties.MaxLength = 100
        Me.teSupplierName.Size = New System.Drawing.Size(94, 20)
        Me.teSupplierName.StyleController = Me.LayoutControl1
        Me.teSupplierName.TabIndex = 5
        ConditionValidationRule2.ConditionOperator = DevExpress.XtraEditors.DXErrorProvider.ConditionOperator.IsNotBlank
        ConditionValidationRule2.ErrorText = "Can not be Blank"
        Me.dxvpSuppliers.SetValidationRule(Me.teSupplierName, ConditionValidationRule2)
        '
        'LayoutControlGroup1
        '
        Me.LayoutControlGroup1.CustomizationFormText = "LayoutControlGroup1"
        Me.LayoutControlGroup1.Items.AddRange(New DevExpress.XtraLayout.BaseLayoutItem() {Me.LayoutControlItem3, Me.LayoutControlItem8, Me.LayoutControlItem6, Me.LayoutControlItem9, Me.LayoutControlItem10, Me.EmptySpaceItem2, Me.LayoutControlItem7, Me.EmptySpaceItem3, Me.LayoutControlGroup2, Me.LayoutControlItem12, Me.EmptySpaceItem1, Me.EmptySpaceItem4, Me.LayoutControlItem5, Me.LayoutControlItem4, Me.EmptySpaceItem5, Me.EmptySpaceItem6, Me.EmptySpaceItem7})
        Me.LayoutControlGroup1.Location = New System.Drawing.Point(0, 0)
        Me.LayoutControlGroup1.Name = "LayoutControlGroup1"
        Me.LayoutControlGroup1.Size = New System.Drawing.Size(633, 315)
        Me.LayoutControlGroup1.Spacing = New DevExpress.XtraLayout.Utils.Padding(0, 0, 0, 0)
        Me.LayoutControlGroup1.Text = "LayoutControlGroup1"
        Me.LayoutControlGroup1.TextVisible = False
        '
        'LayoutControlItem3
        '
        Me.LayoutControlItem3.Control = Me.teAddressLine1
        Me.LayoutControlItem3.CustomizationFormText = "Address"
        Me.LayoutControlItem3.Location = New System.Drawing.Point(0, 117)
        Me.LayoutControlItem3.Name = "LayoutControlItem3"
        Me.LayoutControlItem3.Size = New System.Drawing.Size(315, 31)
        Me.LayoutControlItem3.Text = "Address Line 1"
        Me.LayoutControlItem3.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem3.TextSize = New System.Drawing.Size(70, 20)
        '
        'LayoutControlItem8
        '
        Me.LayoutControlItem8.Control = Me.lblCustomerID
        Me.LayoutControlItem8.CustomizationFormText = "LayoutControlItem8"
        Me.LayoutControlItem8.Location = New System.Drawing.Point(0, 31)
        Me.LayoutControlItem8.Name = "LayoutControlItem8"
        Me.LayoutControlItem8.ShowInCustomizationForm = False
        Me.LayoutControlItem8.Size = New System.Drawing.Size(631, 24)
        Me.LayoutControlItem8.Text = "LayoutControlItem8"
        Me.LayoutControlItem8.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem8.TextSize = New System.Drawing.Size(0, 0)
        Me.LayoutControlItem8.TextToControlDistance = 0
        Me.LayoutControlItem8.TextVisible = False
        Me.LayoutControlItem8.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never
        '
        'LayoutControlItem6
        '
        Me.LayoutControlItem6.Control = Me.teAccountNo
        Me.LayoutControlItem6.CustomizationFormText = "Account No"
        Me.LayoutControlItem6.Location = New System.Drawing.Point(0, 241)
        Me.LayoutControlItem6.Name = "LayoutControlItem6"
        Me.LayoutControlItem6.Size = New System.Drawing.Size(315, 72)
        Me.LayoutControlItem6.Text = "Account No"
        Me.LayoutControlItem6.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem6.TextSize = New System.Drawing.Size(70, 20)
        '
        'LayoutControlItem9
        '
        Me.LayoutControlItem9.Control = Me.teAddressLine2
        Me.LayoutControlItem9.CustomizationFormText = "Address"
        Me.LayoutControlItem9.Location = New System.Drawing.Point(0, 148)
        Me.LayoutControlItem9.Name = "LayoutControlItem9"
        Me.LayoutControlItem9.Size = New System.Drawing.Size(315, 31)
        Me.LayoutControlItem9.Text = "Address Line 2"
        Me.LayoutControlItem9.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem9.TextSize = New System.Drawing.Size(70, 20)
        '
        'LayoutControlItem10
        '
        Me.LayoutControlItem10.Control = Me.teAddressLine3
        Me.LayoutControlItem10.CustomizationFormText = "Address"
        Me.LayoutControlItem10.Location = New System.Drawing.Point(0, 179)
        Me.LayoutControlItem10.Name = "LayoutControlItem10"
        Me.LayoutControlItem10.Size = New System.Drawing.Size(315, 31)
        Me.LayoutControlItem10.Text = "Address Line 3"
        Me.LayoutControlItem10.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem10.TextSize = New System.Drawing.Size(70, 20)
        '
        'EmptySpaceItem2
        '
        Me.EmptySpaceItem2.CustomizationFormText = "EmptySpaceItem2"
        Me.EmptySpaceItem2.Location = New System.Drawing.Point(0, 0)
        Me.EmptySpaceItem2.MaxSize = New System.Drawing.Size(0, 31)
        Me.EmptySpaceItem2.MinSize = New System.Drawing.Size(10, 10)
        Me.EmptySpaceItem2.Name = "EmptySpaceItem2"
        Me.EmptySpaceItem2.Size = New System.Drawing.Size(631, 31)
        Me.EmptySpaceItem2.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom
        Me.EmptySpaceItem2.Text = "EmptySpaceItem2"
        Me.EmptySpaceItem2.TextSize = New System.Drawing.Size(0, 0)
        '
        'LayoutControlItem7
        '
        Me.LayoutControlItem7.Control = Me.ceBank
        Me.LayoutControlItem7.CustomizationFormText = "Bank"
        Me.LayoutControlItem7.Location = New System.Drawing.Point(315, 241)
        Me.LayoutControlItem7.Name = "LayoutControlItem7"
        Me.LayoutControlItem7.Size = New System.Drawing.Size(316, 72)
        Me.LayoutControlItem7.Text = "Bank"
        Me.LayoutControlItem7.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem7.TextSize = New System.Drawing.Size(70, 20)
        '
        'EmptySpaceItem3
        '
        Me.EmptySpaceItem3.CustomizationFormText = "EmptySpaceItem3"
        Me.EmptySpaceItem3.Location = New System.Drawing.Point(316, 86)
        Me.EmptySpaceItem3.Name = "EmptySpaceItem3"
        Me.EmptySpaceItem3.Size = New System.Drawing.Size(315, 31)
        Me.EmptySpaceItem3.Text = "EmptySpaceItem3"
        Me.EmptySpaceItem3.TextSize = New System.Drawing.Size(0, 0)
        '
        'LayoutControlGroup2
        '
        Me.LayoutControlGroup2.CustomizationFormText = "LayoutControlGroup2"
        Me.LayoutControlGroup2.GroupBordersVisible = False
        Me.LayoutControlGroup2.Items.AddRange(New DevExpress.XtraLayout.BaseLayoutItem() {Me.LayoutControlItem2, Me.LayoutControlItem11})
        Me.LayoutControlGroup2.Location = New System.Drawing.Point(0, 86)
        Me.LayoutControlGroup2.Name = "LayoutControlGroup2"
        Me.LayoutControlGroup2.Size = New System.Drawing.Size(316, 31)
        Me.LayoutControlGroup2.Text = "LayoutControlGroup2"
        Me.LayoutControlGroup2.TextVisible = False
        '
        'LayoutControlItem2
        '
        Me.LayoutControlItem2.Control = Me.teSupplierName
        Me.LayoutControlItem2.CustomizationFormText = "Customer Name"
        Me.LayoutControlItem2.Location = New System.Drawing.Point(136, 0)
        Me.LayoutControlItem2.Name = "LayoutControlItem2"
        Me.LayoutControlItem2.Size = New System.Drawing.Size(180, 31)
        Me.LayoutControlItem2.Text = "Vendor Name"
        Me.LayoutControlItem2.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem2.TextSize = New System.Drawing.Size(70, 0)
        '
        'LayoutControlItem11
        '
        Me.LayoutControlItem11.Control = Me.cbeSalutation
        Me.LayoutControlItem11.CustomizationFormText = "Supplier Name"
        Me.LayoutControlItem11.Location = New System.Drawing.Point(0, 0)
        Me.LayoutControlItem11.MaxSize = New System.Drawing.Size(136, 31)
        Me.LayoutControlItem11.MinSize = New System.Drawing.Size(136, 31)
        Me.LayoutControlItem11.Name = "LayoutControlItem11"
        Me.LayoutControlItem11.Size = New System.Drawing.Size(136, 31)
        Me.LayoutControlItem11.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom
        Me.LayoutControlItem11.Text = "Supplier Name"
        Me.LayoutControlItem11.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem11.TextSize = New System.Drawing.Size(70, 20)
        Me.LayoutControlItem11.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never
        '
        'LayoutControlItem12
        '
        Me.LayoutControlItem12.Control = Me.seSupplierNo
        Me.LayoutControlItem12.CustomizationFormText = "Supplier No"
        Me.LayoutControlItem12.Location = New System.Drawing.Point(0, 55)
        Me.LayoutControlItem12.Name = "LayoutControlItem12"
        Me.LayoutControlItem12.Size = New System.Drawing.Size(157, 31)
        Me.LayoutControlItem12.Text = "Vendor No"
        Me.LayoutControlItem12.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem12.TextSize = New System.Drawing.Size(70, 20)
        '
        'EmptySpaceItem1
        '
        Me.EmptySpaceItem1.CustomizationFormText = "EmptySpaceItem1"
        Me.EmptySpaceItem1.Location = New System.Drawing.Point(315, 55)
        Me.EmptySpaceItem1.Name = "EmptySpaceItem1"
        Me.EmptySpaceItem1.Size = New System.Drawing.Size(316, 31)
        Me.EmptySpaceItem1.Text = "EmptySpaceItem1"
        Me.EmptySpaceItem1.TextSize = New System.Drawing.Size(0, 0)
        '
        'EmptySpaceItem4
        '
        Me.EmptySpaceItem4.CustomizationFormText = "EmptySpaceItem4"
        Me.EmptySpaceItem4.Location = New System.Drawing.Point(157, 55)
        Me.EmptySpaceItem4.Name = "EmptySpaceItem4"
        Me.EmptySpaceItem4.Size = New System.Drawing.Size(158, 31)
        Me.EmptySpaceItem4.Text = "EmptySpaceItem4"
        Me.EmptySpaceItem4.TextSize = New System.Drawing.Size(0, 0)
        '
        'LayoutControlItem5
        '
        Me.LayoutControlItem5.Control = Me.teTelephone
        Me.LayoutControlItem5.CustomizationFormText = "Telephone No"
        Me.LayoutControlItem5.Location = New System.Drawing.Point(0, 210)
        Me.LayoutControlItem5.Name = "LayoutControlItem5"
        Me.LayoutControlItem5.Size = New System.Drawing.Size(315, 31)
        Me.LayoutControlItem5.Text = "Telephone No"
        Me.LayoutControlItem5.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem5.TextSize = New System.Drawing.Size(70, 20)
        '
        'LayoutControlItem4
        '
        Me.LayoutControlItem4.Control = Me.teEstate
        Me.LayoutControlItem4.CustomizationFormText = "Estate"
        Me.LayoutControlItem4.Location = New System.Drawing.Point(315, 210)
        Me.LayoutControlItem4.Name = "LayoutControlItem4"
        Me.LayoutControlItem4.Size = New System.Drawing.Size(316, 31)
        Me.LayoutControlItem4.Text = "Fax"
        Me.LayoutControlItem4.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem4.TextSize = New System.Drawing.Size(70, 20)
        '
        'EmptySpaceItem5
        '
        Me.EmptySpaceItem5.CustomizationFormText = "EmptySpaceItem5"
        Me.EmptySpaceItem5.Location = New System.Drawing.Point(315, 117)
        Me.EmptySpaceItem5.Name = "EmptySpaceItem5"
        Me.EmptySpaceItem5.Size = New System.Drawing.Size(316, 31)
        Me.EmptySpaceItem5.Text = "EmptySpaceItem5"
        Me.EmptySpaceItem5.TextSize = New System.Drawing.Size(0, 0)
        '
        'EmptySpaceItem6
        '
        Me.EmptySpaceItem6.CustomizationFormText = "EmptySpaceItem6"
        Me.EmptySpaceItem6.Location = New System.Drawing.Point(315, 148)
        Me.EmptySpaceItem6.Name = "EmptySpaceItem6"
        Me.EmptySpaceItem6.Size = New System.Drawing.Size(316, 31)
        Me.EmptySpaceItem6.Text = "EmptySpaceItem6"
        Me.EmptySpaceItem6.TextSize = New System.Drawing.Size(0, 0)
        '
        'EmptySpaceItem7
        '
        Me.EmptySpaceItem7.CustomizationFormText = "EmptySpaceItem7"
        Me.EmptySpaceItem7.Location = New System.Drawing.Point(315, 179)
        Me.EmptySpaceItem7.Name = "EmptySpaceItem7"
        Me.EmptySpaceItem7.Size = New System.Drawing.Size(316, 31)
        Me.EmptySpaceItem7.Text = "EmptySpaceItem7"
        Me.EmptySpaceItem7.TextSize = New System.Drawing.Size(0, 0)
        '
        'XtraTabPage2
        '
        Me.XtraTabPage2.Controls.Add(Me.gcSuppliers)
        Me.XtraTabPage2.Name = "XtraTabPage2"
        Me.XtraTabPage2.Size = New System.Drawing.Size(633, 314)
        Me.XtraTabPage2.Text = "History Data"
        '
        'gcSuppliers
        '
        Me.gcSuppliers.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gcSuppliers.Location = New System.Drawing.Point(0, 0)
        Me.gcSuppliers.MainView = Me.gvSuppliers
        Me.gcSuppliers.Name = "gcSuppliers"
        Me.gcSuppliers.Size = New System.Drawing.Size(633, 314)
        Me.gcSuppliers.TabIndex = 0
        Me.gcSuppliers.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.gvSuppliers, Me.GridView2})
        '
        'gvSuppliers
        '
        Me.gvSuppliers.Columns.AddRange(New DevExpress.XtraGrid.Columns.GridColumn() {Me.GridColumn1, Me.GridColumn2, Me.GridColumn3, Me.GridColumn4, Me.GridColumn5, Me.GridColumn6, Me.GridColumn7, Me.GridColumn8, Me.GridColumn9, Me.GridColumn10, Me.GridColumn11, Me.GridColumn12})
        Me.gvSuppliers.GridControl = Me.gcSuppliers
        Me.gvSuppliers.Name = "gvSuppliers"
        Me.gvSuppliers.OptionsBehavior.Editable = False
        Me.gvSuppliers.OptionsView.ColumnAutoWidth = False
        Me.gvSuppliers.OptionsView.EnableAppearanceOddRow = True
        Me.gvSuppliers.OptionsView.ShowAutoFilterRow = True
        Me.gvSuppliers.OptionsView.ShowFooter = True
        '
        'GridColumn1
        '
        Me.GridColumn1.Caption = "SupplierID"
        Me.GridColumn1.FieldName = "SupplierID"
        Me.GridColumn1.Name = "GridColumn1"
        '
        'GridColumn2
        '
        Me.GridColumn2.Caption = "Vendor No"
        Me.GridColumn2.FieldName = "SupplierNo"
        Me.GridColumn2.Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left
        Me.GridColumn2.Name = "GridColumn2"
        Me.GridColumn2.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Count
        Me.GridColumn2.Visible = True
        Me.GridColumn2.VisibleIndex = 0
        Me.GridColumn2.Width = 100
        '
        'GridColumn3
        '
        Me.GridColumn3.Caption = "Vendor Name"
        Me.GridColumn3.FieldName = "SupplierName"
        Me.GridColumn3.Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left
        Me.GridColumn3.Name = "GridColumn3"
        Me.GridColumn3.Visible = True
        Me.GridColumn3.VisibleIndex = 1
        Me.GridColumn3.Width = 150
        '
        'GridColumn4
        '
        Me.GridColumn4.Caption = "Address Line1"
        Me.GridColumn4.FieldName = "AddressLine1"
        Me.GridColumn4.Name = "GridColumn4"
        Me.GridColumn4.Visible = True
        Me.GridColumn4.VisibleIndex = 2
        Me.GridColumn4.Width = 100
        '
        'GridColumn5
        '
        Me.GridColumn5.Caption = "Telephone"
        Me.GridColumn5.FieldName = "Telephone"
        Me.GridColumn5.Name = "GridColumn5"
        Me.GridColumn5.Visible = True
        Me.GridColumn5.VisibleIndex = 3
        Me.GridColumn5.Width = 100
        '
        'GridColumn6
        '
        Me.GridColumn6.Caption = "Fax"
        Me.GridColumn6.FieldName = "Estate"
        Me.GridColumn6.Name = "GridColumn6"
        Me.GridColumn6.Visible = True
        Me.GridColumn6.VisibleIndex = 4
        Me.GridColumn6.Width = 100
        '
        'GridColumn7
        '
        Me.GridColumn7.Caption = "Account No"
        Me.GridColumn7.FieldName = "AccountNo"
        Me.GridColumn7.Name = "GridColumn7"
        Me.GridColumn7.Visible = True
        Me.GridColumn7.VisibleIndex = 5
        Me.GridColumn7.Width = 100
        '
        'GridColumn8
        '
        Me.GridColumn8.Caption = "Bank"
        Me.GridColumn8.FieldName = "Bank"
        Me.GridColumn8.Name = "GridColumn8"
        Me.GridColumn8.Visible = True
        Me.GridColumn8.VisibleIndex = 6
        Me.GridColumn8.Width = 100
        '
        'GridColumn9
        '
        Me.GridColumn9.Caption = "Created By"
        Me.GridColumn9.FieldName = "CreatedBy"
        Me.GridColumn9.Name = "GridColumn9"
        Me.GridColumn9.Visible = True
        Me.GridColumn9.VisibleIndex = 7
        Me.GridColumn9.Width = 120
        '
        'GridColumn10
        '
        Me.GridColumn10.Caption = "Created Date"
        Me.GridColumn10.DisplayFormat.FormatString = "dd-MMM-yy h:mm tt"
        Me.GridColumn10.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime
        Me.GridColumn10.FieldName = "CreatedDate"
        Me.GridColumn10.FilterMode = DevExpress.XtraGrid.ColumnFilterMode.DisplayText
        Me.GridColumn10.Name = "GridColumn10"
        Me.GridColumn10.Visible = True
        Me.GridColumn10.VisibleIndex = 8
        Me.GridColumn10.Width = 120
        '
        'GridColumn11
        '
        Me.GridColumn11.Caption = "Updated By"
        Me.GridColumn11.FieldName = "UpdatedBy"
        Me.GridColumn11.Name = "GridColumn11"
        Me.GridColumn11.Visible = True
        Me.GridColumn11.VisibleIndex = 9
        Me.GridColumn11.Width = 120
        '
        'GridColumn12
        '
        Me.GridColumn12.Caption = "Updated Date"
        Me.GridColumn12.DisplayFormat.FormatString = "dd-MMM-yy h:mm tt"
        Me.GridColumn12.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime
        Me.GridColumn12.FieldName = "UpdatedDate"
        Me.GridColumn12.FilterMode = DevExpress.XtraGrid.ColumnFilterMode.DisplayText
        Me.GridColumn12.Name = "GridColumn12"
        Me.GridColumn12.Visible = True
        Me.GridColumn12.VisibleIndex = 10
        Me.GridColumn12.Width = 120
        '
        'GridView2
        '
        Me.GridView2.GridControl = Me.gcSuppliers
        Me.GridView2.Name = "GridView2"
        '
        'frmSuppliers
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(642, 391)
        Me.Controls.Add(Me.xTab1)
        Me.Controls.Add(Me.barDockControlLeft)
        Me.Controls.Add(Me.barDockControlRight)
        Me.Controls.Add(Me.barDockControlBottom)
        Me.Controls.Add(Me.barDockControlTop)
        Me.KeyPreview = True
        Me.Name = "frmSuppliers"
        Me.Text = "Vendors"
        CType(Me.BarManager1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.xTab1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.xTab1.ResumeLayout(False)
        Me.XtraTabPage1.ResumeLayout(False)
        CType(Me.LayoutControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.LayoutControl1.ResumeLayout(False)
        CType(Me.seSupplierNo.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cbeSalutation.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ceBank.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.teAddressLine3.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.teAddressLine2.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.teAccountNo.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.teTelephone.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.teEstate.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.teAddressLine1.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.teSupplierName.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlGroup1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem8, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem9, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem10, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.EmptySpaceItem2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.EmptySpaceItem3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlGroup2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem11, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem12, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.EmptySpaceItem1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.EmptySpaceItem4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.EmptySpaceItem5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.EmptySpaceItem6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.EmptySpaceItem7, System.ComponentModel.ISupportInitialize).EndInit()
        Me.XtraTabPage2.ResumeLayout(False)
        CType(Me.gcSuppliers, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gvSuppliers, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridView2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dxvpSuppliers, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents BarManager1 As DevExpress.XtraBars.BarManager
    Friend WithEvents Bar2 As DevExpress.XtraBars.Bar
    Friend WithEvents bbSave As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents bbNew As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents bbDelete As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents bbDisplaySelected As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents bbRefresh As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents bbClose As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents barDockControlTop As DevExpress.XtraBars.BarDockControl
    Friend WithEvents barDockControlBottom As DevExpress.XtraBars.BarDockControl
    Friend WithEvents barDockControlLeft As DevExpress.XtraBars.BarDockControl
    Friend WithEvents barDockControlRight As DevExpress.XtraBars.BarDockControl
    Friend WithEvents xTab1 As DevExpress.XtraTab.XtraTabControl
    Friend WithEvents XtraTabPage1 As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents XtraTabPage2 As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents gcSuppliers As DevExpress.XtraGrid.GridControl
    Friend WithEvents gvSuppliers As DevExpress.XtraGrid.Views.Grid.GridView
    Friend WithEvents GridView2 As DevExpress.XtraGrid.Views.Grid.GridView
    Friend WithEvents LayoutControl1 As DevExpress.XtraLayout.LayoutControl
    Friend WithEvents LayoutControlGroup1 As DevExpress.XtraLayout.LayoutControlGroup
    Friend WithEvents teSupplierName As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LayoutControlItem2 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents teAddressLine1 As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LayoutControlItem3 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents teEstate As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LayoutControlItem4 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents teTelephone As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LayoutControlItem5 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents teAccountNo As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LayoutControlItem6 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents lblCustomerID As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LayoutControlItem8 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents teAddressLine3 As DevExpress.XtraEditors.TextEdit
    Friend WithEvents teAddressLine2 As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LayoutControlItem9 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents LayoutControlItem10 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents GridColumn1 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn2 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn3 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn4 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn5 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn6 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn7 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn8 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents dxvpSuppliers As DevExpress.XtraEditors.DXErrorProvider.DXValidationProvider
    Friend WithEvents EmptySpaceItem2 As DevExpress.XtraLayout.EmptySpaceItem
    Friend WithEvents ceBank As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents LayoutControlItem7 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents cbeSalutation As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents LayoutControlItem11 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents EmptySpaceItem3 As DevExpress.XtraLayout.EmptySpaceItem
    Friend WithEvents LayoutControlGroup2 As DevExpress.XtraLayout.LayoutControlGroup
    Friend WithEvents seSupplierNo As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents LayoutControlItem12 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents EmptySpaceItem1 As DevExpress.XtraLayout.EmptySpaceItem
    Friend WithEvents bbPrint As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents GridColumn9 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn10 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn11 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn12 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents EmptySpaceItem4 As DevExpress.XtraLayout.EmptySpaceItem
    Friend WithEvents EmptySpaceItem5 As DevExpress.XtraLayout.EmptySpaceItem
    Friend WithEvents EmptySpaceItem6 As DevExpress.XtraLayout.EmptySpaceItem
    Friend WithEvents EmptySpaceItem7 As DevExpress.XtraLayout.EmptySpaceItem
End Class
