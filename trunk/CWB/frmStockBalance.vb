Imports CWBCommon.CWBStock
Imports CWBCommon.CWBSuppliers
Imports CWBCommon.CWBConstants
Imports DevExpress.XtraGrid
Imports DevExpress.XtraEditors.Controls
Imports DevExpress.XtraEditors
Imports DevExpress.XtraEditors.ViewInfo

Public Class frmStockBalance


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
    Private Sub frmStockBalance_Activated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Activated

    End Sub

    Private Sub frmStockBalance_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.Chr(27) Then
            Me.Close()
        End If
    End Sub
#End Region

#Region "Populate Grid"
    Public Sub PopulateGrid()
        Try
            gcStockBalance.DataSource = CWBStock.GetAllStockItems().Tables(0)
        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Print Preview"
    Private Sub sbPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sbPrint.Click
        PrintPreview(gcStockBalance, "Stock Balance Report")
    End Sub
#End Region

#Region "Radio Group Events"
    Private Sub rgShowBy_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rgShowBy.SelectedIndexChanged


        Try
            Dim SelectedValue As Object = rgShowBy.EditValue

            Select Case SelectedValue.ToString()
                Case 1
                    Me.gvStockBalance.ClearGrouping()
                    Me.gvStockBalance.ClearColumnsFilter()
                    Me.PopulateGrid()
                    gvStockBalance.ClearColumnsFilter()
                    gvStockBalance.Columns("StockType").FilterInfo = New DevExpress.XtraGrid.Columns.ColumnFilterInfo(Columns.ColumnFilterType.Custom, Nothing, "[StockType] LIKE 'STOCK ITEM%'", "")
                    GridColumn5.Visible = True
                    GridColumn6.Visible = True
                    GridColumn7.Visible = True
                    GridColumn9.Visible = True

                Case 2
                    Me.gvStockBalance.ClearGrouping()
                    Me.gvStockBalance.ClearColumnsFilter()
                    Me.PopulateGrid()
                    gvStockBalance.ClearColumnsFilter()
                    gvStockBalance.Columns("StockType").FilterInfo = New DevExpress.XtraGrid.Columns.ColumnFilterInfo(Columns.ColumnFilterType.Custom, Nothing, "[StockType] LIKE 'SERVICE%'", "")
                    GridColumn5.Visible = False
                    GridColumn6.Visible = False
                    GridColumn7.Visible = False
                    GridColumn9.Visible = False


                    'Case 3
                    '    gvStockBalance.ClearColumnsFilter()
                    '    gvStockBalance.ClearGrouping()



            End Select

        Catch ex As Exception
            MessageError(ex.ToString)
        End Try














    End Sub
#End Region

  
    
End Class