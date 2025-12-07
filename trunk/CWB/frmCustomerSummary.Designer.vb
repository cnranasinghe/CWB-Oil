<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCustomerSummary
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
        Me.LayoutControl1 = New DevExpress.XtraLayout.LayoutControl
        Me.sbSaveDiscount = New DevExpress.XtraEditors.SimpleButton
        Me.sbPrint = New DevExpress.XtraEditors.SimpleButton
        Me.sbProcess = New DevExpress.XtraEditors.SimpleButton
        Me.gcCustomerSummary = New DevExpress.XtraGrid.GridControl
        Me.gvCustomerSummary = New DevExpress.XtraGrid.Views.Grid.GridView
        Me.GridColumn1 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn2 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn3 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn8 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn4 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn5 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.RepositoryItemSpinEdit1 = New DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit
        Me.GridColumn6 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn7 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.GridColumn9 = New DevExpress.XtraGrid.Columns.GridColumn
        Me.teVehicleNo = New DevExpress.XtraEditors.TextEdit
        Me.deToDate = New DevExpress.XtraEditors.DateEdit
        Me.deFromDate = New DevExpress.XtraEditors.DateEdit
        Me.leCustomers = New DevExpress.XtraEditors.LookUpEdit
        Me.LayoutControlGroup1 = New DevExpress.XtraLayout.LayoutControlGroup
        Me.LayoutControlItem5 = New DevExpress.XtraLayout.LayoutControlItem
        Me.EmptySpaceItem1 = New DevExpress.XtraLayout.EmptySpaceItem
        Me.EmptySpaceItem2 = New DevExpress.XtraLayout.EmptySpaceItem
        Me.LayoutControlGroup2 = New DevExpress.XtraLayout.LayoutControlGroup
        Me.LayoutControlItem2 = New DevExpress.XtraLayout.LayoutControlItem
        Me.LayoutControlItem3 = New DevExpress.XtraLayout.LayoutControlItem
        Me.LayoutControlItem1 = New DevExpress.XtraLayout.LayoutControlItem
        Me.LayoutControlItem4 = New DevExpress.XtraLayout.LayoutControlItem
        Me.LayoutControlItem6 = New DevExpress.XtraLayout.LayoutControlItem
        Me.LayoutControlItem7 = New DevExpress.XtraLayout.LayoutControlItem
        Me.EmptySpaceItem3 = New DevExpress.XtraLayout.EmptySpaceItem
        Me.LayoutControlItem8 = New DevExpress.XtraLayout.LayoutControlItem
        CType(Me.LayoutControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.LayoutControl1.SuspendLayout()
        CType(Me.gcCustomerSummary, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gvCustomerSummary, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemSpinEdit1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.teVehicleNo.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.deToDate.Properties.VistaTimeProperties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.deToDate.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.deFromDate.Properties.VistaTimeProperties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.deFromDate.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.leCustomers.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlGroup1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.EmptySpaceItem1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.EmptySpaceItem2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlGroup2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.EmptySpaceItem3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem8, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'LayoutControl1
        '
        Me.LayoutControl1.Controls.Add(Me.sbSaveDiscount)
        Me.LayoutControl1.Controls.Add(Me.sbPrint)
        Me.LayoutControl1.Controls.Add(Me.sbProcess)
        Me.LayoutControl1.Controls.Add(Me.gcCustomerSummary)
        Me.LayoutControl1.Controls.Add(Me.teVehicleNo)
        Me.LayoutControl1.Controls.Add(Me.deToDate)
        Me.LayoutControl1.Controls.Add(Me.deFromDate)
        Me.LayoutControl1.Controls.Add(Me.leCustomers)
        Me.LayoutControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LayoutControl1.Location = New System.Drawing.Point(0, 0)
        Me.LayoutControl1.Name = "LayoutControl1"
        Me.LayoutControl1.Root = Me.LayoutControlGroup1
        Me.LayoutControl1.Size = New System.Drawing.Size(931, 496)
        Me.LayoutControl1.TabIndex = 0
        Me.LayoutControl1.Text = "LayoutControl1"
        '
        'sbSaveDiscount
        '
        Me.sbSaveDiscount.Location = New System.Drawing.Point(533, 463)
        Me.sbSaveDiscount.Name = "sbSaveDiscount"
        Me.sbSaveDiscount.Size = New System.Drawing.Size(195, 30)
        Me.sbSaveDiscount.StyleController = Me.LayoutControl1
        Me.sbSaveDiscount.TabIndex = 11
        Me.sbSaveDiscount.Text = "Save Discount"
        '
        'sbPrint
        '
        Me.sbPrint.Location = New System.Drawing.Point(733, 463)
        Me.sbPrint.Name = "sbPrint"
        Me.sbPrint.Size = New System.Drawing.Size(195, 30)
        Me.sbPrint.StyleController = Me.LayoutControl1
        Me.sbPrint.TabIndex = 10
        Me.sbPrint.Text = "Print"
        '
        'sbProcess
        '
        Me.sbProcess.Location = New System.Drawing.Point(560, 86)
        Me.sbProcess.Name = "sbProcess"
        Me.sbProcess.Size = New System.Drawing.Size(156, 30)
        Me.sbProcess.StyleController = Me.LayoutControl1
        Me.sbProcess.TabIndex = 4
        Me.sbProcess.Text = "Process"
        '
        'gcCustomerSummary
        '
        Me.gcCustomerSummary.EmbeddedNavigator.Buttons.Append.Visible = False
        Me.gcCustomerSummary.EmbeddedNavigator.Buttons.CancelEdit.Visible = False
        Me.gcCustomerSummary.EmbeddedNavigator.Buttons.Edit.Visible = False
        Me.gcCustomerSummary.EmbeddedNavigator.Buttons.EndEdit.Visible = False
        Me.gcCustomerSummary.EmbeddedNavigator.Buttons.First.Visible = False
        Me.gcCustomerSummary.EmbeddedNavigator.Buttons.Last.Visible = False
        Me.gcCustomerSummary.EmbeddedNavigator.Buttons.Next.Visible = False
        Me.gcCustomerSummary.EmbeddedNavigator.Buttons.NextPage.Visible = False
        Me.gcCustomerSummary.EmbeddedNavigator.Buttons.Prev.Visible = False
        Me.gcCustomerSummary.EmbeddedNavigator.Buttons.PrevPage.Visible = False
        Me.gcCustomerSummary.Location = New System.Drawing.Point(7, 126)
        Me.gcCustomerSummary.MainView = Me.gvCustomerSummary
        Me.gcCustomerSummary.Name = "gcCustomerSummary"
        Me.gcCustomerSummary.RepositoryItems.AddRange(New DevExpress.XtraEditors.Repository.RepositoryItem() {Me.RepositoryItemSpinEdit1})
        Me.gcCustomerSummary.Size = New System.Drawing.Size(918, 329)
        Me.gcCustomerSummary.TabIndex = 8
        Me.gcCustomerSummary.UseEmbeddedNavigator = True
        Me.gcCustomerSummary.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.gvCustomerSummary})
        '
        'gvCustomerSummary
        '
        Me.gvCustomerSummary.Appearance.Row.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gvCustomerSummary.Appearance.Row.Options.UseFont = True
        Me.gvCustomerSummary.AppearancePrint.FooterPanel.BackColor = System.Drawing.Color.White
        Me.gvCustomerSummary.AppearancePrint.FooterPanel.BackColor2 = System.Drawing.Color.White
        Me.gvCustomerSummary.AppearancePrint.FooterPanel.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gvCustomerSummary.AppearancePrint.FooterPanel.ForeColor = System.Drawing.Color.Black
        Me.gvCustomerSummary.AppearancePrint.FooterPanel.Options.UseBackColor = True
        Me.gvCustomerSummary.AppearancePrint.FooterPanel.Options.UseFont = True
        Me.gvCustomerSummary.AppearancePrint.FooterPanel.Options.UseForeColor = True
        Me.gvCustomerSummary.AppearancePrint.HeaderPanel.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gvCustomerSummary.AppearancePrint.HeaderPanel.Options.UseFont = True
        Me.gvCustomerSummary.AppearancePrint.Row.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gvCustomerSummary.AppearancePrint.Row.Options.UseFont = True
        Me.gvCustomerSummary.Columns.AddRange(New DevExpress.XtraGrid.Columns.GridColumn() {Me.GridColumn1, Me.GridColumn2, Me.GridColumn3, Me.GridColumn8, Me.GridColumn4, Me.GridColumn5, Me.GridColumn6, Me.GridColumn7, Me.GridColumn9})
        Me.gvCustomerSummary.GridControl = Me.gcCustomerSummary
        Me.gvCustomerSummary.Name = "gvCustomerSummary"
        Me.gvCustomerSummary.OptionsPrint.UsePrintStyles = True
        Me.gvCustomerSummary.OptionsView.ShowFooter = True
        Me.gvCustomerSummary.OptionsView.ShowGroupPanel = False
        '
        'GridColumn1
        '
        Me.GridColumn1.Caption = "Sales No"
        Me.GridColumn1.FieldName = "SalesNo"
        Me.GridColumn1.Name = "GridColumn1"
        Me.GridColumn1.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Count
        Me.GridColumn1.Visible = True
        Me.GridColumn1.VisibleIndex = 0
        '
        'GridColumn2
        '
        Me.GridColumn2.Caption = "Customer"
        Me.GridColumn2.FieldName = "CustomerName"
        Me.GridColumn2.Name = "GridColumn2"
        Me.GridColumn2.OptionsColumn.AllowEdit = False
        Me.GridColumn2.Visible = True
        Me.GridColumn2.VisibleIndex = 1
        '
        'GridColumn3
        '
        Me.GridColumn3.Caption = "Date"
        Me.GridColumn3.DisplayFormat.FormatString = "dd-MMM-yy"
        Me.GridColumn3.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime
        Me.GridColumn3.FieldName = "SalesDate"
        Me.GridColumn3.Name = "GridColumn3"
        Me.GridColumn3.OptionsColumn.AllowEdit = False
        Me.GridColumn3.Visible = True
        Me.GridColumn3.VisibleIndex = 2
        '
        'GridColumn8
        '
        Me.GridColumn8.Caption = "Vehicle No"
        Me.GridColumn8.FieldName = "ReferenceNo"
        Me.GridColumn8.Name = "GridColumn8"
        Me.GridColumn8.OptionsColumn.AllowEdit = False
        Me.GridColumn8.Visible = True
        Me.GridColumn8.VisibleIndex = 3
        '
        'GridColumn4
        '
        Me.GridColumn4.Caption = "Total"
        Me.GridColumn4.DisplayFormat.FormatString = "{0:n2}"
        Me.GridColumn4.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.GridColumn4.FieldName = "Total"
        Me.GridColumn4.Name = "GridColumn4"
        Me.GridColumn4.OptionsColumn.AllowEdit = False
        Me.GridColumn4.SummaryItem.DisplayFormat = "{0:n2}"
        Me.GridColumn4.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum
        Me.GridColumn4.Visible = True
        Me.GridColumn4.VisibleIndex = 4
        '
        'GridColumn5
        '
        Me.GridColumn5.Caption = "Discount (%)"
        Me.GridColumn5.ColumnEdit = Me.RepositoryItemSpinEdit1
        Me.GridColumn5.DisplayFormat.FormatString = "{0:n2}"
        Me.GridColumn5.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.GridColumn5.FieldName = "Percentage"
        Me.GridColumn5.Name = "GridColumn5"
        Me.GridColumn5.Visible = True
        Me.GridColumn5.VisibleIndex = 5
        '
        'RepositoryItemSpinEdit1
        '
        Me.RepositoryItemSpinEdit1.AutoHeight = False
        Me.RepositoryItemSpinEdit1.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton})
        Me.RepositoryItemSpinEdit1.Name = "RepositoryItemSpinEdit1"
        '
        'GridColumn6
        '
        Me.GridColumn6.Caption = "Discount Amount"
        Me.GridColumn6.DisplayFormat.FormatString = "{0:n2}"
        Me.GridColumn6.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.GridColumn6.FieldName = "Discount"
        Me.GridColumn6.Name = "GridColumn6"
        Me.GridColumn6.OptionsColumn.AllowEdit = False
        Me.GridColumn6.SummaryItem.DisplayFormat = "{0:n2}"
        Me.GridColumn6.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum
        Me.GridColumn6.Visible = True
        Me.GridColumn6.VisibleIndex = 6
        '
        'GridColumn7
        '
        Me.GridColumn7.Caption = "Grand Total"
        Me.GridColumn7.DisplayFormat.FormatString = "{0:n2}"
        Me.GridColumn7.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.GridColumn7.FieldName = "GrandTotal"
        Me.GridColumn7.Name = "GridColumn7"
        Me.GridColumn7.OptionsColumn.AllowEdit = False
        Me.GridColumn7.SummaryItem.DisplayFormat = "{0:n2}"
        Me.GridColumn7.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum
        Me.GridColumn7.Visible = True
        Me.GridColumn7.VisibleIndex = 7
        '
        'GridColumn9
        '
        Me.GridColumn9.Caption = "SalesID"
        Me.GridColumn9.FieldName = "SalesID"
        Me.GridColumn9.Name = "GridColumn9"
        '
        'teVehicleNo
        '
        Me.teVehicleNo.EnterMoveNextControl = True
        Me.teVehicleNo.Location = New System.Drawing.Point(274, 90)
        Me.teVehicleNo.Name = "teVehicleNo"
        Me.teVehicleNo.Size = New System.Drawing.Size(279, 20)
        Me.teVehicleNo.StyleController = Me.LayoutControl1
        Me.teVehicleNo.TabIndex = 3
        '
        'deToDate
        '
        Me.deToDate.EditValue = Nothing
        Me.deToDate.EnterMoveNextControl = True
        Me.deToDate.Location = New System.Drawing.Point(525, 28)
        Me.deToDate.Name = "deToDate"
        Me.deToDate.Properties.AllowNullInput = DevExpress.Utils.DefaultBoolean.[False]
        Me.deToDate.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.deToDate.Properties.DisplayFormat.FormatString = "dd-MMM-yy"
        Me.deToDate.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime
        Me.deToDate.Properties.EditFormat.FormatString = "dd-MMM-yy"
        Me.deToDate.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.DateTime
        Me.deToDate.Properties.Mask.EditMask = "dd-MMM-yy"
        Me.deToDate.Properties.VistaTimeProperties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton})
        Me.deToDate.Size = New System.Drawing.Size(187, 20)
        Me.deToDate.StyleController = Me.LayoutControl1
        Me.deToDate.TabIndex = 1
        '
        'deFromDate
        '
        Me.deFromDate.EditValue = Nothing
        Me.deFromDate.EnterMoveNextControl = True
        Me.deFromDate.Location = New System.Drawing.Point(274, 28)
        Me.deFromDate.Name = "deFromDate"
        Me.deFromDate.Properties.AllowNullInput = DevExpress.Utils.DefaultBoolean.[False]
        Me.deFromDate.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.deFromDate.Properties.DisplayFormat.FormatString = "dd-MMM-yy"
        Me.deFromDate.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime
        Me.deFromDate.Properties.EditFormat.FormatString = "dd-MMM-yy"
        Me.deFromDate.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.DateTime
        Me.deFromDate.Properties.Mask.EditMask = "dd-MMM-yy"
        Me.deFromDate.Properties.VistaTimeProperties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton})
        Me.deFromDate.Size = New System.Drawing.Size(186, 20)
        Me.deFromDate.StyleController = Me.LayoutControl1
        Me.deFromDate.TabIndex = 0
        '
        'leCustomers
        '
        Me.leCustomers.EnterMoveNextControl = True
        Me.leCustomers.Location = New System.Drawing.Point(274, 59)
        Me.leCustomers.Name = "leCustomers"
        Me.leCustomers.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo), New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Glyph, "Clear", -1, True, True, False, DevExpress.XtraEditors.ImageLocation.MiddleCenter, Nothing)})
        Me.leCustomers.Properties.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.leCustomers.Properties.Columns.AddRange(New DevExpress.XtraEditors.Controls.LookUpColumnInfo() {New DevExpress.XtraEditors.Controls.LookUpColumnInfo("CustomerNo", "Cutomer No", 20, DevExpress.Utils.FormatType.None, "", True, DevExpress.Utils.HorzAlignment.[Default], DevExpress.Data.ColumnSortOrder.None), New DevExpress.XtraEditors.Controls.LookUpColumnInfo("CustomerName", "Customner Name", 20, DevExpress.Utils.FormatType.None, "", True, DevExpress.Utils.HorzAlignment.[Default], DevExpress.Data.ColumnSortOrder.None)})
        Me.leCustomers.Properties.NullText = ""
        Me.leCustomers.Size = New System.Drawing.Size(438, 20)
        Me.leCustomers.StyleController = Me.LayoutControl1
        Me.leCustomers.TabIndex = 2
        '
        'LayoutControlGroup1
        '
        Me.LayoutControlGroup1.CustomizationFormText = "LayoutControlGroup1"
        Me.LayoutControlGroup1.Items.AddRange(New DevExpress.XtraLayout.BaseLayoutItem() {Me.LayoutControlItem5, Me.EmptySpaceItem1, Me.EmptySpaceItem2, Me.LayoutControlGroup2, Me.LayoutControlItem7, Me.EmptySpaceItem3, Me.LayoutControlItem8})
        Me.LayoutControlGroup1.Location = New System.Drawing.Point(0, 0)
        Me.LayoutControlGroup1.Name = "LayoutControlGroup1"
        Me.LayoutControlGroup1.Size = New System.Drawing.Size(931, 496)
        Me.LayoutControlGroup1.Spacing = New DevExpress.XtraLayout.Utils.Padding(0, 0, 0, 0)
        Me.LayoutControlGroup1.Text = "LayoutControlGroup1"
        Me.LayoutControlGroup1.TextVisible = False
        '
        'LayoutControlItem5
        '
        Me.LayoutControlItem5.Control = Me.gcCustomerSummary
        Me.LayoutControlItem5.CustomizationFormText = "LayoutControlItem5"
        Me.LayoutControlItem5.Location = New System.Drawing.Point(0, 119)
        Me.LayoutControlItem5.Name = "LayoutControlItem5"
        Me.LayoutControlItem5.Size = New System.Drawing.Size(929, 340)
        Me.LayoutControlItem5.Text = "LayoutControlItem5"
        Me.LayoutControlItem5.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem5.TextSize = New System.Drawing.Size(0, 0)
        Me.LayoutControlItem5.TextToControlDistance = 0
        Me.LayoutControlItem5.TextVisible = False
        '
        'EmptySpaceItem1
        '
        Me.EmptySpaceItem1.CustomizationFormText = "EmptySpaceItem1"
        Me.EmptySpaceItem1.Location = New System.Drawing.Point(0, 0)
        Me.EmptySpaceItem1.Name = "EmptySpaceItem1"
        Me.EmptySpaceItem1.Size = New System.Drawing.Size(210, 119)
        Me.EmptySpaceItem1.Text = "EmptySpaceItem1"
        Me.EmptySpaceItem1.TextSize = New System.Drawing.Size(0, 0)
        '
        'EmptySpaceItem2
        '
        Me.EmptySpaceItem2.CustomizationFormText = "EmptySpaceItem2"
        Me.EmptySpaceItem2.Location = New System.Drawing.Point(719, 0)
        Me.EmptySpaceItem2.Name = "EmptySpaceItem2"
        Me.EmptySpaceItem2.Size = New System.Drawing.Size(210, 119)
        Me.EmptySpaceItem2.Text = "EmptySpaceItem2"
        Me.EmptySpaceItem2.TextSize = New System.Drawing.Size(0, 0)
        '
        'LayoutControlGroup2
        '
        Me.LayoutControlGroup2.CustomizationFormText = "LayoutControlGroup2"
        Me.LayoutControlGroup2.Items.AddRange(New DevExpress.XtraLayout.BaseLayoutItem() {Me.LayoutControlItem2, Me.LayoutControlItem3, Me.LayoutControlItem1, Me.LayoutControlItem4, Me.LayoutControlItem6})
        Me.LayoutControlGroup2.Location = New System.Drawing.Point(210, 0)
        Me.LayoutControlGroup2.Name = "LayoutControlGroup2"
        Me.LayoutControlGroup2.Size = New System.Drawing.Size(509, 119)
        Me.LayoutControlGroup2.Text = "Select Date Range to Extract Customer Summary"
        '
        'LayoutControlItem2
        '
        Me.LayoutControlItem2.Control = Me.deFromDate
        Me.LayoutControlItem2.CustomizationFormText = "From"
        Me.LayoutControlItem2.Location = New System.Drawing.Point(0, 0)
        Me.LayoutControlItem2.Name = "LayoutControlItem2"
        Me.LayoutControlItem2.Size = New System.Drawing.Size(251, 31)
        Me.LayoutControlItem2.Text = "From"
        Me.LayoutControlItem2.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem2.TextSize = New System.Drawing.Size(49, 20)
        '
        'LayoutControlItem3
        '
        Me.LayoutControlItem3.Control = Me.deToDate
        Me.LayoutControlItem3.CustomizationFormText = "To"
        Me.LayoutControlItem3.Location = New System.Drawing.Point(251, 0)
        Me.LayoutControlItem3.Name = "LayoutControlItem3"
        Me.LayoutControlItem3.Size = New System.Drawing.Size(252, 31)
        Me.LayoutControlItem3.Text = "To"
        Me.LayoutControlItem3.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem3.TextSize = New System.Drawing.Size(49, 20)
        '
        'LayoutControlItem1
        '
        Me.LayoutControlItem1.Control = Me.leCustomers
        Me.LayoutControlItem1.CustomizationFormText = "Customer"
        Me.LayoutControlItem1.Location = New System.Drawing.Point(0, 31)
        Me.LayoutControlItem1.Name = "LayoutControlItem1"
        Me.LayoutControlItem1.Size = New System.Drawing.Size(503, 31)
        Me.LayoutControlItem1.Text = "Customer"
        Me.LayoutControlItem1.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem1.TextSize = New System.Drawing.Size(49, 20)
        '
        'LayoutControlItem4
        '
        Me.LayoutControlItem4.Control = Me.teVehicleNo
        Me.LayoutControlItem4.CustomizationFormText = "Vehicle No"
        Me.LayoutControlItem4.Location = New System.Drawing.Point(0, 62)
        Me.LayoutControlItem4.Name = "LayoutControlItem4"
        Me.LayoutControlItem4.Size = New System.Drawing.Size(344, 33)
        Me.LayoutControlItem4.Text = "Vehicle No"
        Me.LayoutControlItem4.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem4.TextSize = New System.Drawing.Size(49, 20)
        '
        'LayoutControlItem6
        '
        Me.LayoutControlItem6.Control = Me.sbProcess
        Me.LayoutControlItem6.CustomizationFormText = "LayoutControlItem6"
        Me.LayoutControlItem6.Location = New System.Drawing.Point(344, 62)
        Me.LayoutControlItem6.MaxSize = New System.Drawing.Size(0, 33)
        Me.LayoutControlItem6.MinSize = New System.Drawing.Size(61, 33)
        Me.LayoutControlItem6.Name = "LayoutControlItem6"
        Me.LayoutControlItem6.Padding = New DevExpress.XtraLayout.Utils.Padding(1, 1, 1, 1)
        Me.LayoutControlItem6.Size = New System.Drawing.Size(159, 33)
        Me.LayoutControlItem6.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom
        Me.LayoutControlItem6.Text = "LayoutControlItem6"
        Me.LayoutControlItem6.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem6.TextSize = New System.Drawing.Size(0, 0)
        Me.LayoutControlItem6.TextToControlDistance = 0
        Me.LayoutControlItem6.TextVisible = False
        '
        'LayoutControlItem7
        '
        Me.LayoutControlItem7.Control = Me.sbPrint
        Me.LayoutControlItem7.CustomizationFormText = "LayoutControlItem7"
        Me.LayoutControlItem7.Location = New System.Drawing.Point(729, 459)
        Me.LayoutControlItem7.MaxSize = New System.Drawing.Size(200, 35)
        Me.LayoutControlItem7.MinSize = New System.Drawing.Size(200, 35)
        Me.LayoutControlItem7.Name = "LayoutControlItem7"
        Me.LayoutControlItem7.Padding = New DevExpress.XtraLayout.Utils.Padding(2, 2, 2, 2)
        Me.LayoutControlItem7.Size = New System.Drawing.Size(200, 35)
        Me.LayoutControlItem7.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom
        Me.LayoutControlItem7.Text = "LayoutControlItem7"
        Me.LayoutControlItem7.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem7.TextSize = New System.Drawing.Size(0, 0)
        Me.LayoutControlItem7.TextToControlDistance = 0
        Me.LayoutControlItem7.TextVisible = False
        '
        'EmptySpaceItem3
        '
        Me.EmptySpaceItem3.CustomizationFormText = "EmptySpaceItem3"
        Me.EmptySpaceItem3.Location = New System.Drawing.Point(0, 459)
        Me.EmptySpaceItem3.Name = "EmptySpaceItem3"
        Me.EmptySpaceItem3.Size = New System.Drawing.Size(529, 35)
        Me.EmptySpaceItem3.Text = "EmptySpaceItem3"
        Me.EmptySpaceItem3.TextSize = New System.Drawing.Size(0, 0)
        '
        'LayoutControlItem8
        '
        Me.LayoutControlItem8.Control = Me.sbSaveDiscount
        Me.LayoutControlItem8.CustomizationFormText = "LayoutControlItem8"
        Me.LayoutControlItem8.Location = New System.Drawing.Point(529, 459)
        Me.LayoutControlItem8.MaxSize = New System.Drawing.Size(200, 35)
        Me.LayoutControlItem8.MinSize = New System.Drawing.Size(200, 35)
        Me.LayoutControlItem8.Name = "LayoutControlItem8"
        Me.LayoutControlItem8.Padding = New DevExpress.XtraLayout.Utils.Padding(2, 2, 2, 2)
        Me.LayoutControlItem8.Size = New System.Drawing.Size(200, 35)
        Me.LayoutControlItem8.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom
        Me.LayoutControlItem8.Text = "LayoutControlItem8"
        Me.LayoutControlItem8.TextLocation = DevExpress.Utils.Locations.Left
        Me.LayoutControlItem8.TextSize = New System.Drawing.Size(0, 0)
        Me.LayoutControlItem8.TextToControlDistance = 0
        Me.LayoutControlItem8.TextVisible = False
        '
        'frmCustomerSummary
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(931, 496)
        Me.Controls.Add(Me.LayoutControl1)
        Me.KeyPreview = True
        Me.Name = "frmCustomerSummary"
        Me.Text = "Customer Summary"
        CType(Me.LayoutControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.LayoutControl1.ResumeLayout(False)
        CType(Me.gcCustomerSummary, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gvCustomerSummary, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemSpinEdit1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.teVehicleNo.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.deToDate.Properties.VistaTimeProperties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.deToDate.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.deFromDate.Properties.VistaTimeProperties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.deFromDate.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.leCustomers.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlGroup1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.EmptySpaceItem1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.EmptySpaceItem2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlGroup2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.EmptySpaceItem3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem8, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents LayoutControl1 As DevExpress.XtraLayout.LayoutControl
    Friend WithEvents LayoutControlGroup1 As DevExpress.XtraLayout.LayoutControlGroup
    Friend WithEvents leCustomers As DevExpress.XtraEditors.LookUpEdit
    Friend WithEvents LayoutControlItem1 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents deFromDate As DevExpress.XtraEditors.DateEdit
    Friend WithEvents LayoutControlItem2 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents deToDate As DevExpress.XtraEditors.DateEdit
    Friend WithEvents LayoutControlItem3 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents teVehicleNo As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LayoutControlItem4 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents gcCustomerSummary As DevExpress.XtraGrid.GridControl
    Friend WithEvents gvCustomerSummary As DevExpress.XtraGrid.Views.Grid.GridView
    Friend WithEvents LayoutControlItem5 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents EmptySpaceItem1 As DevExpress.XtraLayout.EmptySpaceItem
    Friend WithEvents EmptySpaceItem2 As DevExpress.XtraLayout.EmptySpaceItem
    Friend WithEvents LayoutControlGroup2 As DevExpress.XtraLayout.LayoutControlGroup
    Friend WithEvents sbProcess As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents LayoutControlItem6 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents GridColumn1 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn2 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn3 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn4 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn5 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn6 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn7 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents sbPrint As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents LayoutControlItem7 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents EmptySpaceItem3 As DevExpress.XtraLayout.EmptySpaceItem
    Friend WithEvents GridColumn8 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents RepositoryItemSpinEdit1 As DevExpress.XtraEditors.Repository.RepositoryItemSpinEdit
    Friend WithEvents sbSaveDiscount As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents LayoutControlItem8 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents GridColumn9 As DevExpress.XtraGrid.Columns.GridColumn
End Class
