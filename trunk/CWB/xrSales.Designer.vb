<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Public Class xrSales
    Inherits DevExpress.XtraReports.UI.XtraReport

    'XtraReport overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Designer
    'It can be modified using the Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Detail = New DevExpress.XtraReports.UI.DetailBand
        Me.xrtStock = New DevExpress.XtraReports.UI.XRTable
        Me.XrTableRow2 = New DevExpress.XtraReports.UI.XRTableRow
        Me.XrTableCell5 = New DevExpress.XtraReports.UI.XRTableCell
        Me.XrTableCell1 = New DevExpress.XtraReports.UI.XRTableCell
        Me.XrTableCell2 = New DevExpress.XtraReports.UI.XRTableCell
        Me.XrTableCell6 = New DevExpress.XtraReports.UI.XRTableCell
        Me.XrTableCell4 = New DevExpress.XtraReports.UI.XRTableCell
        Me.PageHeader = New DevExpress.XtraReports.UI.PageHeaderBand
        Me.xrlAddress3 = New DevExpress.XtraReports.UI.XRLabel
        Me.xrlAddress2 = New DevExpress.XtraReports.UI.XRLabel
        Me.xrlAddress1 = New DevExpress.XtraReports.UI.XRLabel
        Me.xrlSalesNo = New DevExpress.XtraReports.UI.XRLabel
        Me.xrlCustomers = New DevExpress.XtraReports.UI.XRLabel
        Me.xrlVehicleNo = New DevExpress.XtraReports.UI.XRLabel
        Me.xrlDate = New DevExpress.XtraReports.UI.XRLabel
        Me.xrlModal = New DevExpress.XtraReports.UI.XRLabel
        Me.xrlKm2 = New DevExpress.XtraReports.UI.XRLabel
        Me.xrlKM1 = New DevExpress.XtraReports.UI.XRLabel
        Me.xrlPaid = New DevExpress.XtraReports.UI.XRLabel
        Me.xrlTotal = New DevExpress.XtraReports.UI.XRLabel
        Me.xrlDue = New DevExpress.XtraReports.UI.XRLabel
        Me.xrlNextReading = New DevExpress.XtraReports.UI.XRLabel
        Me.xrlMeterReading = New DevExpress.XtraReports.UI.XRLabel
        Me.PageFooter = New DevExpress.XtraReports.UI.PageFooterBand
        Me.xrlDiscount = New DevExpress.XtraReports.UI.XRLabel
        Me.xrlMessage = New DevExpress.XtraReports.UI.XRLabel
        Me.xrlCh4 = New DevExpress.XtraReports.UI.XRLabel
        Me.xrlCh3 = New DevExpress.XtraReports.UI.XRLabel
        Me.xrlCh2 = New DevExpress.XtraReports.UI.XRLabel
        Me.xrlCh1 = New DevExpress.XtraReports.UI.XRLabel
        Me.BottomMargin = New DevExpress.XtraReports.UI.BottomMarginBand
        CType(Me.xrtStock, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me, System.ComponentModel.ISupportInitialize).BeginInit()
        '
        'Detail
        '
        Me.Detail.Controls.AddRange(New DevExpress.XtraReports.UI.XRControl() {Me.xrtStock})
        Me.Detail.Height = 20
        Me.Detail.Name = "Detail"
        Me.Detail.Padding = New DevExpress.XtraPrinting.PaddingInfo(0, 0, 0, 0, 100.0!)
        Me.Detail.TextAlignment = DevExpress.XtraPrinting.TextAlignment.TopLeft
        '
        'xrtStock
        '
        Me.xrtStock.Borders = DevExpress.XtraPrinting.BorderSide.None
        Me.xrtStock.Location = New System.Drawing.Point(25, 0)
        Me.xrtStock.Name = "xrtStock"
        Me.xrtStock.Rows.AddRange(New DevExpress.XtraReports.UI.XRTableRow() {Me.XrTableRow2})
        Me.xrtStock.Size = New System.Drawing.Size(730, 20)
        Me.xrtStock.StylePriority.UseBorders = False
        '
        'XrTableRow2
        '
        Me.XrTableRow2.Cells.AddRange(New DevExpress.XtraReports.UI.XRTableCell() {Me.XrTableCell5, Me.XrTableCell1, Me.XrTableCell2, Me.XrTableCell6, Me.XrTableCell4})
        Me.XrTableRow2.Name = "XrTableRow2"
        Me.XrTableRow2.Size = New System.Drawing.Size(730, 20)
        '
        'XrTableCell5
        '
        Me.XrTableCell5.CanGrow = False
        Me.XrTableCell5.Location = New System.Drawing.Point(0, 0)
        Me.XrTableCell5.Name = "XrTableCell5"
        Me.XrTableCell5.Padding = New DevExpress.XtraPrinting.PaddingInfo(2, 2, 0, 0, 100.0!)
        Me.XrTableCell5.Size = New System.Drawing.Size(116, 20)
        Me.XrTableCell5.StylePriority.UseTextAlignment = False
        Me.XrTableCell5.TextAlignment = DevExpress.XtraPrinting.TextAlignment.MiddleLeft
        '
        'XrTableCell1
        '
        Me.XrTableCell1.CanGrow = False
        Me.XrTableCell1.Location = New System.Drawing.Point(116, 0)
        Me.XrTableCell1.Name = "XrTableCell1"
        Me.XrTableCell1.Padding = New DevExpress.XtraPrinting.PaddingInfo(2, 2, 0, 0, 100.0!)
        Me.XrTableCell1.Size = New System.Drawing.Size(235, 20)
        Me.XrTableCell1.StylePriority.UseTextAlignment = False
        Me.XrTableCell1.TextAlignment = DevExpress.XtraPrinting.TextAlignment.MiddleLeft
        '
        'XrTableCell2
        '
        Me.XrTableCell2.Location = New System.Drawing.Point(351, 0)
        Me.XrTableCell2.Name = "XrTableCell2"
        Me.XrTableCell2.Padding = New DevExpress.XtraPrinting.PaddingInfo(2, 2, 0, 0, 100.0!)
        Me.XrTableCell2.Size = New System.Drawing.Size(107, 20)
        Me.XrTableCell2.StylePriority.UseTextAlignment = False
        Me.XrTableCell2.TextAlignment = DevExpress.XtraPrinting.TextAlignment.MiddleRight
        '
        'XrTableCell6
        '
        Me.XrTableCell6.Location = New System.Drawing.Point(458, 0)
        Me.XrTableCell6.Name = "XrTableCell6"
        Me.XrTableCell6.Padding = New DevExpress.XtraPrinting.PaddingInfo(2, 2, 0, 0, 100.0!)
        Me.XrTableCell6.Size = New System.Drawing.Size(141, 20)
        Me.XrTableCell6.StylePriority.UseTextAlignment = False
        Me.XrTableCell6.TextAlignment = DevExpress.XtraPrinting.TextAlignment.MiddleCenter
        '
        'XrTableCell4
        '
        Me.XrTableCell4.Location = New System.Drawing.Point(599, 0)
        Me.XrTableCell4.Name = "XrTableCell4"
        Me.XrTableCell4.Padding = New DevExpress.XtraPrinting.PaddingInfo(2, 2, 0, 0, 100.0!)
        Me.XrTableCell4.Size = New System.Drawing.Size(131, 20)
        Me.XrTableCell4.StylePriority.UseTextAlignment = False
        Me.XrTableCell4.TextAlignment = DevExpress.XtraPrinting.TextAlignment.MiddleRight
        '
        'PageHeader
        '
        Me.PageHeader.Controls.AddRange(New DevExpress.XtraReports.UI.XRControl() {Me.xrlAddress1, Me.xrlAddress2, Me.xrlAddress3, Me.xrlSalesNo, Me.xrlModal, Me.xrlVehicleNo, Me.xrlDate, Me.xrlCustomers})
        Me.PageHeader.Height = 154
        Me.PageHeader.Name = "PageHeader"
        Me.PageHeader.Padding = New DevExpress.XtraPrinting.PaddingInfo(0, 0, 0, 0, 100.0!)
        Me.PageHeader.TextAlignment = DevExpress.XtraPrinting.TextAlignment.TopLeft
        '
        'xrlAddress3
        '
        Me.xrlAddress3.Location = New System.Drawing.Point(74, 78)
        Me.xrlAddress3.Multiline = True
        Me.xrlAddress3.Name = "xrlAddress3"
        Me.xrlAddress3.Padding = New DevExpress.XtraPrinting.PaddingInfo(2, 2, 0, 0, 100.0!)
        Me.xrlAddress3.Size = New System.Drawing.Size(442, 20)
        Me.xrlAddress3.StylePriority.UseTextAlignment = False
        Me.xrlAddress3.Text = "" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.xrlAddress3.TextAlignment = DevExpress.XtraPrinting.TextAlignment.TopLeft
        '
        'xrlAddress2
        '
        Me.xrlAddress2.Location = New System.Drawing.Point(74, 58)
        Me.xrlAddress2.Multiline = True
        Me.xrlAddress2.Name = "xrlAddress2"
        Me.xrlAddress2.Padding = New DevExpress.XtraPrinting.PaddingInfo(2, 2, 0, 0, 100.0!)
        Me.xrlAddress2.Size = New System.Drawing.Size(442, 20)
        Me.xrlAddress2.StylePriority.UseTextAlignment = False
        Me.xrlAddress2.Text = "" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.xrlAddress2.TextAlignment = DevExpress.XtraPrinting.TextAlignment.TopLeft
        '
        'xrlAddress1
        '
        Me.xrlAddress1.Location = New System.Drawing.Point(74, 38)
        Me.xrlAddress1.Multiline = True
        Me.xrlAddress1.Name = "xrlAddress1"
        Me.xrlAddress1.Padding = New DevExpress.XtraPrinting.PaddingInfo(2, 2, 0, 0, 100.0!)
        Me.xrlAddress1.Size = New System.Drawing.Size(442, 20)
        Me.xrlAddress1.StylePriority.UseTextAlignment = False
        Me.xrlAddress1.Text = "" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.xrlAddress1.TextAlignment = DevExpress.XtraPrinting.TextAlignment.TopLeft
        '
        'xrlSalesNo
        '
        Me.xrlSalesNo.Location = New System.Drawing.Point(626, 17)
        Me.xrlSalesNo.Multiline = True
        Me.xrlSalesNo.Name = "xrlSalesNo"
        Me.xrlSalesNo.Padding = New DevExpress.XtraPrinting.PaddingInfo(2, 2, 0, 0, 100.0!)
        Me.xrlSalesNo.Size = New System.Drawing.Size(150, 20)
        Me.xrlSalesNo.StylePriority.UseTextAlignment = False
        Me.xrlSalesNo.Text = "" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.xrlSalesNo.TextAlignment = DevExpress.XtraPrinting.TextAlignment.TopLeft
        '
        'xrlCustomers
        '
        Me.xrlCustomers.Location = New System.Drawing.Point(74, 17)
        Me.xrlCustomers.Multiline = True
        Me.xrlCustomers.Name = "xrlCustomers"
        Me.xrlCustomers.Padding = New DevExpress.XtraPrinting.PaddingInfo(2, 2, 0, 0, 100.0!)
        Me.xrlCustomers.Size = New System.Drawing.Size(442, 20)
        Me.xrlCustomers.StylePriority.UseTextAlignment = False
        Me.xrlCustomers.Text = "" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.xrlCustomers.TextAlignment = DevExpress.XtraPrinting.TextAlignment.TopLeft
        '
        'xrlVehicleNo
        '
        Me.xrlVehicleNo.Location = New System.Drawing.Point(627, 65)
        Me.xrlVehicleNo.Multiline = True
        Me.xrlVehicleNo.Name = "xrlVehicleNo"
        Me.xrlVehicleNo.Padding = New DevExpress.XtraPrinting.PaddingInfo(2, 2, 0, 0, 100.0!)
        Me.xrlVehicleNo.Size = New System.Drawing.Size(150, 20)
        Me.xrlVehicleNo.StylePriority.UseTextAlignment = False
        Me.xrlVehicleNo.Text = "" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.xrlVehicleNo.TextAlignment = DevExpress.XtraPrinting.TextAlignment.TopLeft
        '
        'xrlDate
        '
        Me.xrlDate.Location = New System.Drawing.Point(626, 42)
        Me.xrlDate.Multiline = True
        Me.xrlDate.Name = "xrlDate"
        Me.xrlDate.Padding = New DevExpress.XtraPrinting.PaddingInfo(2, 2, 0, 0, 100.0!)
        Me.xrlDate.Size = New System.Drawing.Size(150, 20)
        Me.xrlDate.StylePriority.UseTextAlignment = False
        Me.xrlDate.Text = "" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.xrlDate.TextAlignment = DevExpress.XtraPrinting.TextAlignment.TopLeft
        '
        'xrlModal
        '
        Me.xrlModal.Location = New System.Drawing.Point(626, 88)
        Me.xrlModal.Name = "xrlModal"
        Me.xrlModal.Padding = New DevExpress.XtraPrinting.PaddingInfo(2, 2, 0, 0, 100.0!)
        Me.xrlModal.Size = New System.Drawing.Size(150, 20)
        Me.xrlModal.StylePriority.UseTextAlignment = False
        Me.xrlModal.TextAlignment = DevExpress.XtraPrinting.TextAlignment.TopLeft
        '
        'xrlKm2
        '
        Me.xrlKm2.Borders = DevExpress.XtraPrinting.BorderSide.None
        Me.xrlKm2.Location = New System.Drawing.Point(185, 74)
        Me.xrlKm2.LockedInUserDesigner = True
        Me.xrlKm2.Name = "xrlKm2"
        Me.xrlKm2.Padding = New DevExpress.XtraPrinting.PaddingInfo(2, 2, 0, 0, 100.0!)
        Me.xrlKm2.Size = New System.Drawing.Size(50, 25)
        Me.xrlKm2.StylePriority.UseBorders = False
        Me.xrlKm2.StylePriority.UseTextAlignment = False
        Me.xrlKm2.TextAlignment = DevExpress.XtraPrinting.TextAlignment.MiddleLeft
        '
        'xrlKM1
        '
        Me.xrlKM1.Borders = DevExpress.XtraPrinting.BorderSide.None
        Me.xrlKM1.Location = New System.Drawing.Point(185, 41)
        Me.xrlKM1.LockedInUserDesigner = True
        Me.xrlKM1.Name = "xrlKM1"
        Me.xrlKM1.Padding = New DevExpress.XtraPrinting.PaddingInfo(2, 2, 0, 0, 100.0!)
        Me.xrlKM1.Size = New System.Drawing.Size(50, 25)
        Me.xrlKM1.StylePriority.UseBorders = False
        Me.xrlKM1.StylePriority.UseTextAlignment = False
        Me.xrlKM1.TextAlignment = DevExpress.XtraPrinting.TextAlignment.TopLeft
        '
        'xrlPaid
        '
        Me.xrlPaid.Borders = DevExpress.XtraPrinting.BorderSide.None
        Me.xrlPaid.Location = New System.Drawing.Point(592, 58)
        Me.xrlPaid.LockedInUserDesigner = True
        Me.xrlPaid.Name = "xrlPaid"
        Me.xrlPaid.Padding = New DevExpress.XtraPrinting.PaddingInfo(2, 2, 0, 0, 100.0!)
        Me.xrlPaid.Size = New System.Drawing.Size(150, 20)
        Me.xrlPaid.StylePriority.UseBorders = False
        Me.xrlPaid.StylePriority.UseTextAlignment = False
        Me.xrlPaid.TextAlignment = DevExpress.XtraPrinting.TextAlignment.TopRight
        '
        'xrlTotal
        '
        Me.xrlTotal.Borders = DevExpress.XtraPrinting.BorderSide.None
        Me.xrlTotal.Location = New System.Drawing.Point(592, 0)
        Me.xrlTotal.LockedInUserDesigner = True
        Me.xrlTotal.Name = "xrlTotal"
        Me.xrlTotal.Padding = New DevExpress.XtraPrinting.PaddingInfo(2, 2, 0, 0, 100.0!)
        Me.xrlTotal.Size = New System.Drawing.Size(150, 20)
        Me.xrlTotal.StylePriority.UseBorders = False
        Me.xrlTotal.StylePriority.UseTextAlignment = False
        Me.xrlTotal.TextAlignment = DevExpress.XtraPrinting.TextAlignment.MiddleRight
        '
        'xrlDue
        '
        Me.xrlDue.Borders = DevExpress.XtraPrinting.BorderSide.None
        Me.xrlDue.Location = New System.Drawing.Point(592, 83)
        Me.xrlDue.LockedInUserDesigner = True
        Me.xrlDue.Name = "xrlDue"
        Me.xrlDue.Padding = New DevExpress.XtraPrinting.PaddingInfo(2, 2, 0, 0, 100.0!)
        Me.xrlDue.Size = New System.Drawing.Size(150, 20)
        Me.xrlDue.StylePriority.UseBorders = False
        Me.xrlDue.StylePriority.UseTextAlignment = False
        Me.xrlDue.TextAlignment = DevExpress.XtraPrinting.TextAlignment.TopRight
        '
        'xrlNextReading
        '
        Me.xrlNextReading.Borders = DevExpress.XtraPrinting.BorderSide.None
        Me.xrlNextReading.Location = New System.Drawing.Point(27, 74)
        Me.xrlNextReading.LockedInUserDesigner = True
        Me.xrlNextReading.Name = "xrlNextReading"
        Me.xrlNextReading.Padding = New DevExpress.XtraPrinting.PaddingInfo(2, 2, 0, 0, 100.0!)
        Me.xrlNextReading.Size = New System.Drawing.Size(150, 25)
        Me.xrlNextReading.StylePriority.UseBorders = False
        Me.xrlNextReading.StylePriority.UseTextAlignment = False
        Me.xrlNextReading.TextAlignment = DevExpress.XtraPrinting.TextAlignment.MiddleRight
        '
        'xrlMeterReading
        '
        Me.xrlMeterReading.Borders = DevExpress.XtraPrinting.BorderSide.None
        Me.xrlMeterReading.Location = New System.Drawing.Point(27, 41)
        Me.xrlMeterReading.LockedInUserDesigner = True
        Me.xrlMeterReading.Name = "xrlMeterReading"
        Me.xrlMeterReading.Padding = New DevExpress.XtraPrinting.PaddingInfo(2, 2, 0, 0, 100.0!)
        Me.xrlMeterReading.Size = New System.Drawing.Size(150, 25)
        Me.xrlMeterReading.StylePriority.UseBorders = False
        Me.xrlMeterReading.StylePriority.UseTextAlignment = False
        Me.xrlMeterReading.TextAlignment = DevExpress.XtraPrinting.TextAlignment.TopRight
        '
        'PageFooter
        '
        Me.PageFooter.Height = 0
        Me.PageFooter.Name = "PageFooter"
        Me.PageFooter.Padding = New DevExpress.XtraPrinting.PaddingInfo(0, 0, 0, 0, 100.0!)
        Me.PageFooter.TextAlignment = DevExpress.XtraPrinting.TextAlignment.TopLeft
        '
        'xrlDiscount
        '
        Me.xrlDiscount.Borders = DevExpress.XtraPrinting.BorderSide.None
        Me.xrlDiscount.Location = New System.Drawing.Point(592, 32)
        Me.xrlDiscount.LockedInUserDesigner = True
        Me.xrlDiscount.Name = "xrlDiscount"
        Me.xrlDiscount.Padding = New DevExpress.XtraPrinting.PaddingInfo(2, 2, 0, 0, 100.0!)
        Me.xrlDiscount.Size = New System.Drawing.Size(150, 20)
        Me.xrlDiscount.StylePriority.UseBorders = False
        Me.xrlDiscount.StylePriority.UseTextAlignment = False
        Me.xrlDiscount.TextAlignment = DevExpress.XtraPrinting.TextAlignment.TopRight
        '
        'xrlMessage
        '
        Me.xrlMessage.Borders = DevExpress.XtraPrinting.BorderSide.None
        Me.xrlMessage.Location = New System.Drawing.Point(25, 0)
        Me.xrlMessage.LockedInUserDesigner = True
        Me.xrlMessage.Name = "xrlMessage"
        Me.xrlMessage.Padding = New DevExpress.XtraPrinting.PaddingInfo(2, 2, 0, 0, 100.0!)
        Me.xrlMessage.Size = New System.Drawing.Size(567, 25)
        Me.xrlMessage.StylePriority.UseBorders = False
        Me.xrlMessage.StylePriority.UseTextAlignment = False
        Me.xrlMessage.TextAlignment = DevExpress.XtraPrinting.TextAlignment.MiddleCenter
        '
        'xrlCh4
        '
        Me.xrlCh4.Borders = DevExpress.XtraPrinting.BorderSide.None
        Me.xrlCh4.Location = New System.Drawing.Point(305, 74)
        Me.xrlCh4.LockedInUserDesigner = True
        Me.xrlCh4.Name = "xrlCh4"
        Me.xrlCh4.Padding = New DevExpress.XtraPrinting.PaddingInfo(2, 2, 0, 0, 100.0!)
        Me.xrlCh4.Size = New System.Drawing.Size(50, 25)
        Me.xrlCh4.StylePriority.UseBorders = False
        Me.xrlCh4.StylePriority.UseTextAlignment = False
        Me.xrlCh4.TextAlignment = DevExpress.XtraPrinting.TextAlignment.MiddleLeft
        '
        'xrlCh3
        '
        Me.xrlCh3.Borders = DevExpress.XtraPrinting.BorderSide.None
        Me.xrlCh3.Location = New System.Drawing.Point(252, 74)
        Me.xrlCh3.LockedInUserDesigner = True
        Me.xrlCh3.Name = "xrlCh3"
        Me.xrlCh3.Padding = New DevExpress.XtraPrinting.PaddingInfo(2, 2, 0, 0, 100.0!)
        Me.xrlCh3.Size = New System.Drawing.Size(50, 25)
        Me.xrlCh3.StylePriority.UseBorders = False
        Me.xrlCh3.StylePriority.UseTextAlignment = False
        Me.xrlCh3.TextAlignment = DevExpress.XtraPrinting.TextAlignment.MiddleLeft
        '
        'xrlCh2
        '
        Me.xrlCh2.Borders = DevExpress.XtraPrinting.BorderSide.None
        Me.xrlCh2.Location = New System.Drawing.Point(305, 41)
        Me.xrlCh2.LockedInUserDesigner = True
        Me.xrlCh2.Name = "xrlCh2"
        Me.xrlCh2.Padding = New DevExpress.XtraPrinting.PaddingInfo(2, 2, 0, 0, 100.0!)
        Me.xrlCh2.Size = New System.Drawing.Size(50, 25)
        Me.xrlCh2.StylePriority.UseBorders = False
        Me.xrlCh2.StylePriority.UseTextAlignment = False
        Me.xrlCh2.TextAlignment = DevExpress.XtraPrinting.TextAlignment.TopLeft
        '
        'xrlCh1
        '
        Me.xrlCh1.Borders = DevExpress.XtraPrinting.BorderSide.None
        Me.xrlCh1.Location = New System.Drawing.Point(252, 41)
        Me.xrlCh1.LockedInUserDesigner = True
        Me.xrlCh1.Name = "xrlCh1"
        Me.xrlCh1.Padding = New DevExpress.XtraPrinting.PaddingInfo(2, 2, 0, 0, 100.0!)
        Me.xrlCh1.Size = New System.Drawing.Size(50, 25)
        Me.xrlCh1.StylePriority.UseBorders = False
        Me.xrlCh1.StylePriority.UseTextAlignment = False
        Me.xrlCh1.TextAlignment = DevExpress.XtraPrinting.TextAlignment.TopLeft
        '
        'BottomMargin
        '
        Me.BottomMargin.Controls.AddRange(New DevExpress.XtraReports.UI.XRControl() {Me.xrlDiscount, Me.xrlMessage, Me.xrlCh4, Me.xrlCh3, Me.xrlCh2, Me.xrlCh1, Me.xrlPaid, Me.xrlTotal, Me.xrlKM1, Me.xrlKm2, Me.xrlNextReading, Me.xrlMeterReading, Me.xrlDue})
        Me.BottomMargin.Height = 122
        Me.BottomMargin.Name = "BottomMargin"
        '
        'xrSales
        '
        Me.Bands.AddRange(New DevExpress.XtraReports.UI.Band() {Me.Detail, Me.PageHeader, Me.PageFooter, Me.BottomMargin})
        Me.Font = New System.Drawing.Font("Tahoma", 8.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margins = New System.Drawing.Printing.Margins(25, 25, 100, 122)
        Me.PageHeight = 590
        Me.PaperKind = System.Drawing.Printing.PaperKind.Custom
        Me.Version = "8.2"
        CType(Me.xrtStock, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub
    Friend WithEvents Detail As DevExpress.XtraReports.UI.DetailBand
    Friend WithEvents PageHeader As DevExpress.XtraReports.UI.PageHeaderBand
    Friend WithEvents PageFooter As DevExpress.XtraReports.UI.PageFooterBand
    Friend WithEvents xrlNextReading As DevExpress.XtraReports.UI.XRLabel
    Friend WithEvents xrlMeterReading As DevExpress.XtraReports.UI.XRLabel
    Friend WithEvents xrlModal As DevExpress.XtraReports.UI.XRLabel
    Friend WithEvents xrtStock As DevExpress.XtraReports.UI.XRTable
    Friend WithEvents XrTableRow2 As DevExpress.XtraReports.UI.XRTableRow
    Friend WithEvents XrTableCell5 As DevExpress.XtraReports.UI.XRTableCell
    Friend WithEvents XrTableCell2 As DevExpress.XtraReports.UI.XRTableCell
    Friend WithEvents XrTableCell6 As DevExpress.XtraReports.UI.XRTableCell
    Friend WithEvents XrTableCell4 As DevExpress.XtraReports.UI.XRTableCell
    Friend WithEvents xrlDue As DevExpress.XtraReports.UI.XRLabel
    Friend WithEvents xrlDate As DevExpress.XtraReports.UI.XRLabel
    Friend WithEvents xrlVehicleNo As DevExpress.XtraReports.UI.XRLabel
    Friend WithEvents xrlCustomers As DevExpress.XtraReports.UI.XRLabel
    Friend WithEvents XrTableCell1 As DevExpress.XtraReports.UI.XRTableCell
    Friend WithEvents xrlTotal As DevExpress.XtraReports.UI.XRLabel
    Friend WithEvents xrlPaid As DevExpress.XtraReports.UI.XRLabel
    Friend WithEvents xrlSalesNo As DevExpress.XtraReports.UI.XRLabel
    Friend WithEvents xrlAddress2 As DevExpress.XtraReports.UI.XRLabel
    Friend WithEvents xrlAddress1 As DevExpress.XtraReports.UI.XRLabel
    Friend WithEvents xrlAddress3 As DevExpress.XtraReports.UI.XRLabel
    Friend WithEvents xrlKM1 As DevExpress.XtraReports.UI.XRLabel
    Friend WithEvents xrlKm2 As DevExpress.XtraReports.UI.XRLabel
    Friend WithEvents xrlCh2 As DevExpress.XtraReports.UI.XRLabel
    Friend WithEvents xrlCh1 As DevExpress.XtraReports.UI.XRLabel
    Friend WithEvents xrlCh4 As DevExpress.XtraReports.UI.XRLabel
    Friend WithEvents xrlCh3 As DevExpress.XtraReports.UI.XRLabel
    Friend WithEvents xrlMessage As DevExpress.XtraReports.UI.XRLabel
    Friend WithEvents xrlDiscount As DevExpress.XtraReports.UI.XRLabel
    Friend WithEvents BottomMargin As DevExpress.XtraReports.UI.BottomMarginBand
End Class
