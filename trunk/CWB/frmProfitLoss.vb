Public Class frmProfitLoss 

#Region "Variables"
    Dim _CWBProfitAndLoss As CWBCommon.CWBProfitAndLoss
#End Region

#Region "Constructors"
    Public ReadOnly Property CWBProfitAndLoss() As CWBCommon.CWBProfitAndLoss
        Get

            If _CWBProfitAndLoss Is Nothing Then
                _CWBProfitAndLoss = New CWBCommon.CWBProfitAndLoss
            End If

            Return _CWBProfitAndLoss
        End Get
    End Property
#End Region

#Region "Form Events"
    Private Sub frmProfitLoss_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.Chr(27) Then
            Me.Close()
        End If
    End Sub

    Private Sub frmProfitLoss_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        deFromDate.EditValue = Date.Today
        deToDate.EditValue = Date.Today
    End Sub
#End Region

#Region "Button Events"

    Private Sub sbProcess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sbProcess.Click
        Try

            If dxvpPnL.Validate Then

                With CWBProfitAndLoss

                    .FromDate = deFromDate.EditValue
                    .ToDate = deToDate.EditValue
                    .GetPLByDates()
                    seStockValue.EditValue = .StockValue
                    seTotalPurchases.EditValue = .TotalPurchases
                    seTotalSales.EditValue = .TotalSales
                    sePurchaseCost.EditValue = .TotalPurchaseCost
                    seTotalExpenses.EditValue = .TotalExpenses

                    seSalesProfit.EditValue = .TotalSales - (.TotalPurchaseCost + .TotalExpenses)
                    seCashtoPay.EditValue = .TotalCashtoPay
                    seCashtoReceive.EditValue = .TotalCashtoReceive
                    sbPrint.Enabled = True
                End With
            Else
                seStockValue.EditValue = 0
                seTotalPurchases.EditValue = 0
                seTotalSales.EditValue = 0
                sePurchaseCost.EditValue = 0
                seTotalExpenses.EditValue = 0

                seSalesProfit.EditValue = 0
                seCashtoPay.EditValue = 0
                seCashtoReceive.EditValue = 0
                sbPrint.Enabled = False
            End If



        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub


    Private Sub sbPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sbPrint.Click
        Try
            Dim _frmPrintPreview As New frmPrint
            Dim _xrProfitnLoss As New xrProfitnLoss




            With _xrProfitnLoss
                .xrlFromTo.Text = "From - " + Format(Me.deFromDate.EditValue, "dd-MMM-yy") + " To -" + Format(Me.deToDate.EditValue, "dd-MMM-yy")
                .xrtcStockValue.Text = FormatNumber(Me.seStockValue.Text, 2, , , )
                .xrtcTotalPurchases.Text = FormatNumber(Me.seTotalPurchases.Text, 2, , , )
                .xrtcTotalSales.Text = FormatNumber(Me.seTotalSales.Text, 2, , , )
                .xrtcTotalPurchasesCost.Text = FormatNumber(Me.sePurchaseCost.Text, 2, , , )
                .xrtcTotalExpenses.Text = FormatNumber(Me.seTotalExpenses.Text, 2, , , )
                .xrtcSalesProfit.Text = FormatNumber(Me.seSalesProfit.Text, 2, , , )
                .xrtcCashttoPay.Text = FormatNumber(Me.seCashtoPay.Text, 2, , , )
                .xrtcCashtoReceive.Text = FormatNumber(Me.seCashtoReceive.Text, 2, , , )

            End With



          



            _frmPrintPreview.PrintControl1.PrintingSystem = _xrProfitnLoss.PrintingSystem
            _xrProfitnLoss.CreateDocument()
            _frmPrintPreview.MdiParent = frmMain
            _frmPrintPreview.Show()
            _frmPrintPreview.BringToFront()
        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
#End Region

  
End Class