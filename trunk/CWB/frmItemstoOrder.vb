Imports CWBCommon.CWBStock
Imports CWBCommon.CWBConstants
Public Class frmItemstoOrder


#Region "Variables"
    Private _CWBStock As CWBCommon.CWBStock
#End Region

#Region "Constructors"
    Public ReadOnly Property CWBSuppliers() As CWBCommon.CWBStock
        Get
            If _CWBStock Is Nothing Then
                _CWBStock = New CWBCommon.CWBStock
            End If

            Return _CWBStock
        End Get
    End Property
#End Region

#Region "Form Events"
    Private Sub frmItemstoOrder_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.Chr(27) Then
            Me.Close()
        End If
    End Sub

    Private Sub frmItemstoOrder_Activated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        Me.PopulateGrid()
    End Sub

#End Region

#Region "Populate Grid"
    Public Sub PopulateGrid()
        Try
            gcItemsToOrder.DataSource = CWBSuppliers.GetAllStockItemsToOrder.Tables(0)
        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Button Events"
    Private Sub sbPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sbPrint.Click
        PrintPreview(gcItemsToOrder, "Items to Order")
    End Sub
#End Region

End Class