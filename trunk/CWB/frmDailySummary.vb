Imports CWBCommon.CWBCollections

Public Class frmDailySummary


#Region "Variables"
    Dim _CWBCollections As CWBCommon.CWBCollections
#End Region

#Region "Constructors"
    Public ReadOnly Property CWBCollections() As CWBCommon.CWBCollections
        Get

            If _CWBCollections Is Nothing Then
                _CWBCollections = New CWBCommon.CWBCollections
            End If

            Return _CWBCollections
        End Get
    End Property
#End Region

#Region "Form Events"
    Private Sub frmCashCollections_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.Chr(27) Then
            Me.Close()
        End If
    End Sub

    Private Sub frmSummary_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.deFromDate.EditValue = Date.Today
        Me.deToDate.EditValue = Date.Today
    End Sub

    Private Sub frmDailySummary_Activated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        Me.PopolateGrid()
    End Sub
#End Region

#Region "Button Events"
    Private Sub sbProcess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sbProcess.Click
        Me.PopolateGrid()
    End Sub
    Private Sub sbPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sbPrint.Click

        Select Case xtab1.SelectedTabPageIndex
            Case 0
                PrintPreview(gcPurchases, "Daily Summary - Purchases")
            Case 1
                PrintPreview(gcSales, "Daily Summary - Sales")
            Case 2
                PrintPreview(gcExpenses, "Daily Summary - Expenses")
            Case 3
                PrintPreview(gcStockItems, "Daily Summary - Stock Items")
            Case 4
                PrintPreview(gcServices, "Daily Summary - Services")

        End Select


    End Sub
#End Region

#Region "Populate Grid"
    Private Sub PopolateGrid()
        If dxvpSalesSummary.Validate Then
            Try
                CWBCollections.FromDate = Me.deFromDate.EditValue
                CWBCollections.ToDate = Me.deToDate.EditValue

                gcPurchases.DataSource = CWBCollections.CollectionGetByDates().Tables(0)

                gcSales.DataSource = CWBCollections.CollectionGetByDates().Tables(1)

                gcExpenses.DataSource = CWBCollections.CollectionGetByDates().Tables(2)


                gcStockItems.DataSource = CWBCollections.CollectionGetByDates().Tables(3)

                gcServices.DataSource = CWBCollections.CollectionGetByDates().Tables(4)


            Catch ex As Exception
                MessageError(ex.ToString)
            End Try

        End If

    End Sub
#End Region

End Class