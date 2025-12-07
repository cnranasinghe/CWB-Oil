
Imports Microsoft.Practices.EnterpriseLibrary.Common
Imports Microsoft.Practices.EnterpriseLibrary.Data
Imports System.Data
Imports System.Data.Common
Imports DevExpress.XtraEditors
Imports DevExpress.XtraPrinting
Imports CWBCommon.CWBConstants

Module CWBWindowsForms

    Public UserName, UserType As String
    Public UserID, EmployerID As Int64




#Region "Form operations"
    Public Sub ShowCWBForm(ByVal frm As DevExpress.XtraEditors.XtraForm)
        frm.MdiParent = frmMain
        frm.Show()
        frm.BringToFront()
    End Sub
#End Region

#Region "Pop Dialog boxes"
    Public Sub MessageOK(ByVal Title As String, ByVal LabelTitle As String, ByVal LabelDes As String)
        Dim f As New frmOk
        f.Text = Title
        f.lblTitle.Text = LabelTitle
        f.lblDescription.Text = LabelDes
        If f.ShowDialog = Windows.Forms.DialogResult.OK Then
            Exit Sub
        End If
    End Sub

    Public Function MessageYesNo(ByVal Title As String, ByVal Label As String, ByVal textbox As String) As DialogResult

        Dim f As New frmUpdateYesNo
        f.Text = Title
        f.lblTitle.Text = Label
        f.lblDescription.Text = textbox
        f.ShowDialog()
        If f.DialogResult = DialogResult.No Then
            Exit Function
        End If


    End Function



    Public Sub MessageSavedOK(ByVal Title As String, ByVal LabelTitle As String, ByVal LabelDes As String)
        Dim f As New frmSavedOk
        f.Text = Title
        f.lblTitle.Text = LabelTitle
        f.lblDescription.Text = LabelDes
        If f.ShowDialog = Windows.Forms.DialogResult.OK Then
            Exit Sub
        End If
    End Sub

#End Region

#Region "Startup form Bar button settings"
    Public Sub SetStartUpBarButton(ByVal bbSave As DevExpress.XtraBars.BarButtonItem, ByVal bbNew As DevExpress.XtraBars.BarButtonItem, ByVal bbDelete As DevExpress.XtraBars.BarButtonItem, ByVal bbClose As DevExpress.XtraBars.BarButtonItem, ByVal bbDisplaySelected As DevExpress.XtraBars.BarButtonItem, ByVal bbRefresh As DevExpress.XtraBars.BarButtonItem, ByVal bbPrint As DevExpress.XtraBars.BarButtonItem)

        bbSave.Caption = "Save"
        bbNew.Caption = "New"
        bbDelete.Caption = "Delete"
        bbClose.Caption = "Close"
        bbDisplaySelected.Caption = "Display Selected"
        bbRefresh.Caption = "Refresh"
        bbPrint.Caption = "Print"

        bbSave.Hint = SAVE_HINT
        bbNew.Hint = NEW_HINT
        bbDelete.Hint = DELETE_HINT
        bbClose.Hint = CLOSE_HINT
        bbDisplaySelected.Hint = DISPLAYSELECTED_HINT
        bbRefresh.Hint = REFRESH_HINT
        bbPrint.Hint = PRINT_HINT

        bbSave.ItemShortcut = New DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.S))
        bbNew.ItemShortcut = New DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.N))
        bbDelete.ItemShortcut = New DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.D))
        bbClose.ItemShortcut = New DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.W))
        bbDisplaySelected.ItemShortcut = New DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.F))
        bbRefresh.ItemShortcut = New DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.R))
        bbPrint.ItemShortcut = New DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.P))


        bbSave.Glyph = CWB.My.Resources.Resources.Save
        bbNew.Glyph = CWB.My.Resources.Resources._New
        bbDelete.Glyph = CWB.My.Resources.Resources.delete
        bbClose.Glyph = CWB.My.Resources.Resources.Close
        bbDisplaySelected.Glyph = CWB.My.Resources.Resources.Show
        bbPrint.Glyph = CWB.My.Resources.Resources.Print
        bbRefresh.Glyph = CWB.My.Resources.Resources.DisplaySelected


        bbSave.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbNew.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbDelete.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbClose.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbDisplaySelected.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbRefresh.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbPrint.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph



    End Sub



#End Region

#Region "Purchases / Sales form Bar button settings"
    Public Sub SetPurchaseSalesBarButton(ByVal bbSave As DevExpress.XtraBars.BarButtonItem, ByVal bbNew As DevExpress.XtraBars.BarButtonItem, ByVal bbClose As DevExpress.XtraBars.BarButtonItem, ByVal bbDisplaySelected As DevExpress.XtraBars.BarButtonItem, ByVal bbRefresh As DevExpress.XtraBars.BarButtonItem, ByVal bbPrint As DevExpress.XtraBars.BarButtonItem)



        bbSave.Caption = "Save"
        bbNew.Caption = "New"
        bbClose.Caption = "Close"
        bbDisplaySelected.Caption = "Display Selected"
        bbRefresh.Caption = "Refresh"
        bbPrint.Caption = "Print"


        bbSave.Hint = SAVE_HINT
        bbNew.Hint = NEW_HINT
        bbClose.Hint = CLOSE_HINT
        bbDisplaySelected.Hint = DISPLAYSELECTED_HINT
        bbRefresh.Hint = REFRESH_HINT
        bbPrint.Hint = PRINT_HINT


        bbSave.Glyph = CWB.My.Resources.Resources.Save
        bbNew.Glyph = CWB.My.Resources.Resources._New
        'bbDelete.Glyph = CWB.My.Resources.Resources.delete
        bbClose.Glyph = CWB.My.Resources.Resources.Close
        bbDisplaySelected.Glyph = CWB.My.Resources.Resources.Show
        bbPrint.Glyph = CWB.My.Resources.Resources.Print
        bbRefresh.Glyph = CWB.My.Resources.Resources.DisplaySelected

        bbSave.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbNew.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbClose.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbDisplaySelected.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbRefresh.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbPrint.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph





    End Sub

    Public Sub SetPurchaseSalesBarButton(ByVal bbSave As DevExpress.XtraBars.BarButtonItem, ByVal bbNew As DevExpress.XtraBars.BarButtonItem, ByVal bbClose As DevExpress.XtraBars.BarButtonItem, ByVal bbDisplaySelected As DevExpress.XtraBars.BarButtonItem, ByVal bbRefresh As DevExpress.XtraBars.BarButtonItem, ByVal bbPrint As DevExpress.XtraBars.BarButtonItem, ByVal bbAutdit As DevExpress.XtraBars.BarButtonItem)



        bbSave.Caption = "Save"
        bbNew.Caption = "New"
        bbClose.Caption = "Close"
        bbDisplaySelected.Caption = "Display Selected"
        bbRefresh.Caption = "Refresh"
        bbPrint.Caption = "Print"
        bbAutdit.Caption = "Audit"


        bbSave.Hint = SAVE_HINT
        bbNew.Hint = NEW_HINT
        bbClose.Hint = CLOSE_HINT
        bbDisplaySelected.Hint = DISPLAYSELECTED_HINT
        bbRefresh.Hint = REFRESH_HINT
        bbPrint.Hint = PRINT_HINT
        bbAutdit.Hint = AUDIT_HINT


        bbSave.Glyph = CWB.My.Resources.Resources.Save
        bbNew.Glyph = CWB.My.Resources.Resources._New
        'bbDelete.Glyph = CWB.My.Resources.Resources.delete
        bbClose.Glyph = CWB.My.Resources.Resources.Close
        bbDisplaySelected.Glyph = CWB.My.Resources.Resources.Show
        bbPrint.Glyph = CWB.My.Resources.Resources.Print
        bbRefresh.Glyph = CWB.My.Resources.Resources.DisplaySelected
        bbAutdit.Glyph = CWB.My.Resources.Resources.StockCat


        bbSave.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbNew.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbClose.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbDisplaySelected.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbRefresh.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbPrint.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbAutdit.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph





    End Sub
#End Region

#Region "Pay Bills Bar Button Settings"
    Public Sub SetPayBillsBarButton(ByVal bbSave As DevExpress.XtraBars.BarButtonItem, ByVal bbPrint As DevExpress.XtraBars.BarButtonItem, ByVal bbShowBy As DevExpress.XtraBars.BarButtonItem, ByVal bbRefresh As DevExpress.XtraBars.BarButtonItem, ByVal bbClose As DevExpress.XtraBars.BarButtonItem, ByVal bbPaid As DevExpress.XtraBars.BarButtonItem)
        bbSave.Hint = SAVE_HINT
        bbPrint.Hint = PRINT_HINT
        bbShowBy.Hint = CE_HINT
        bbRefresh.Hint = REFRESH_HINT
        bbClose.Hint = CLOSE_HINT
        bbPaid.Hint = PAID_HINT


        bbSave.Caption = "Save"
        bbPrint.Caption = "Print"
        bbShowBy.Caption = "Group By Supplier"
        bbRefresh.Caption = "Refresh"
        bbClose.Caption = "Close"
        bbPaid.Caption = "Set As Paid"


        bbSave.Glyph = CWB.My.Resources.Resources.Save
        bbShowBy.Glyph = CWB.My.Resources.Resources.Suppliers
        bbRefresh.Glyph = CWB.My.Resources.Resources.DisplaySelected
        bbPrint.Glyph = CWB.My.Resources.Resources.Print
        bbClose.Glyph = CWB.My.Resources.Resources.Close
        bbPaid.Glyph = CWB.My.Resources.Resources.paid


        bbSave.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbShowBy.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbPrint.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbClose.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbRefresh.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbShowBy.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbPaid.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph


        bbSave.ItemShortcut = New DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.S))


    End Sub
#End Region

#Region "Receive Payments Bar Button Settings"
    Public Sub SetReceivePaymentsBarButton(ByVal bbSave As DevExpress.XtraBars.BarButtonItem, ByVal bbPrint As DevExpress.XtraBars.BarButtonItem, ByVal bbShowBy As DevExpress.XtraBars.BarButtonItem, ByVal bbRefresh As DevExpress.XtraBars.BarButtonItem, ByVal bbClose As DevExpress.XtraBars.BarButtonItem, ByVal bbPaid As DevExpress.XtraBars.BarButtonItem)
        bbSave.Hint = SAVE_HINT
        bbPrint.Hint = PRINT_HINT
        bbShowBy.Hint = CE_HINT
        bbRefresh.Hint = REFRESH_HINT
        bbClose.Hint = CLOSE_HINT
        bbPaid.Hint = PAID_HINT

        bbSave.Caption = "Save"
        bbPrint.Caption = "Print"
        bbShowBy.Caption = "Group By Customer"
        bbRefresh.Caption = "Refresh"
        bbClose.Caption = "Close"
        bbPaid.Caption = "Set As Paid"


        bbSave.Glyph = CWB.My.Resources.Resources.Save
        bbShowBy.Glyph = CWB.My.Resources.Resources.Customers
        bbRefresh.Glyph = CWB.My.Resources.Resources.DisplaySelected
        bbPrint.Glyph = CWB.My.Resources.Resources.Print
        bbClose.Glyph = CWB.My.Resources.Resources.Close
        bbPaid.Glyph = CWB.My.Resources.Resources.Paid


        bbSave.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbShowBy.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbPrint.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbClose.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbRefresh.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbShowBy.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbPaid.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph


        bbSave.ItemShortcut = New DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.S))


    End Sub
#End Region

#Region "Expenses Bar Button Setting"
    Public Sub SetCWBExpenseBarButton(ByVal bbSave As DevExpress.XtraBars.BarButtonItem, ByVal bbNew As DevExpress.XtraBars.BarButtonItem, ByVal bbDelete As DevExpress.XtraBars.BarButtonItem, ByVal bbClose As DevExpress.XtraBars.BarButtonItem, ByVal bbDisplaySelected As DevExpress.XtraBars.BarButtonItem, ByVal bbRefresh As DevExpress.XtraBars.BarButtonItem, ByVal bbPrint As DevExpress.XtraBars.BarButtonItem)

        bbSave.Caption = "Save"
        bbNew.Caption = "New"
        bbDelete.Caption = "Delete"
        bbDisplaySelected.Caption = "Display Selected"
        bbRefresh.Caption = "Refresh"
        bbPrint.Caption = "Print"
        bbClose.Caption = "Close"

        bbSave.Hint = SAVE_HINT
        bbNew.Hint = NEW_HINT
        bbDelete.Hint = DELETE_HINT
        bbClose.Hint = CLOSE_HINT
        bbDisplaySelected.Hint = DISPLAYSELECTED_HINT
        bbRefresh.Hint = REFRESH_HINT
        bbPrint.Hint = PRINT_HINT

        bbSave.ItemShortcut = New DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.S))
        bbNew.ItemShortcut = New DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.N))
        bbDelete.ItemShortcut = New DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.D))
        bbClose.ItemShortcut = New DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.W))
        bbDisplaySelected.ItemShortcut = New DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.F))
        bbRefresh.ItemShortcut = New DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.R))
        bbPrint.ItemShortcut = New DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.P))


        bbSave.Glyph = CWB.My.Resources.Resources.Save
        bbNew.Glyph = CWB.My.Resources.Resources._New
        bbDelete.Glyph = CWB.My.Resources.Resources.delete
        bbClose.Glyph = CWB.My.Resources.Resources.Close
        bbDisplaySelected.Glyph = CWB.My.Resources.Resources.Show
        bbPrint.Glyph = CWB.My.Resources.Resources.Print
        bbRefresh.Glyph = CWB.My.Resources.Resources.DisplaySelected


        bbSave.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbNew.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbDelete.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbClose.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbDisplaySelected.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbRefresh.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbPrint.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph


    End Sub
#End Region

#Region "Summary Bar Button Settings"
    Public Sub SetSummaryBarButton(ByVal bbShowRecords As DevExpress.XtraBars.BarButtonItem, ByVal bbClose As DevExpress.XtraBars.BarButtonItem, ByVal bbRefresh As DevExpress.XtraBars.BarButtonItem)
        bbShowRecords.Caption = "Show Records"
        bbClose.Caption = "Close"
        bbRefresh.Caption = "Refresh"

        bbShowRecords.Hint = SHOWRECORD_HINT
        bbClose.Hint = CLOSE_HINT
        bbRefresh.Hint = REFRESH_HINT

        bbShowRecords.ItemShortcut = New DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.E))
        bbClose.ItemShortcut = New DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.W))
        bbRefresh.ItemShortcut = New DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.R))


        bbShowRecords.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbClose.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        bbRefresh.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph

    End Sub
#End Region

#Region "Print Preview"
    Public Sub PrintPreview(ByVal grid As DevExpress.XtraGrid.GridControl, ByVal reportheadername As String)
        Try
            'Dim Link As New PrintableComponentLink()
            'Link.PrintingSystem = New PrintingSystem
            'Link.Component = grid


            'frmPrint.PrintControl1.PrintingSystem = Link.PrintingSystem

            'Link.Margins = New System.Drawing.Printing.Margins(50, 50, 100, 50)
            'Link.CreateDocument()
            'frmPrint.MdiParent = frmMain
            'frmPrint.Show()
            'frmPrint.BringToFront()

            Dim _CWBCompany As New CWBCommon.CWBCompany
            _CWBCompany.GetCompany()

            Dim Link As New PrintableComponentLink()

            Link.PrintingSystem = New PrintingSystem
            Link.Component = grid

            Dim a As PageHeaderFooter
            a = Link.PageHeaderFooter
            a.Header.LineAlignment = BrickAlignment.Center
            a.Header.Content.Add(" ")
            a.Header.Content.Add("" + _CWBCompany.CompanyName + vbCrLf + reportheadername + vbCrLf + "Printed On -" + Date.Now.ToString("dd-MMM-yy hh:mm:tt"))





            frmPrint.PrintControl1.PrintingSystem = Link.PrintingSystem


            Link.Margins = New System.Drawing.Printing.Margins(50, 50, 100, 50)
            Link.CreateDocument()
            frmPrint.MdiParent = frmMain
            frmPrint.Show()
            frmPrint.BringToFront()



        Catch ex As Exception
            MessageError(ex.ToString)
        End Try


    End Sub
#End Region

#Region "Message Error"
    Public Function MessageError(ByVal strMessage As String) As DialogResult

        Dim f As New frmError
        f.meErrorMessage.Text = strMessage
        f.ShowDialog()
        If f.DialogResult = DialogResult.No Then
            Exit Function
        End If


    End Function
#End Region

End Module
