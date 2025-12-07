Imports CWBCommon.CWBConstants
Imports CWBCommon.CWBCustomers

Public Class frmCustomerSummary

#Region "Variables"
    Dim _CWBCustomers As CWBCommon.CWBCustomers
    Dim _CWBSales As CWBCommon.CWBSales
#End Region

#Region "Constructors"
    Public ReadOnly Property CWBCustomers() As CWBCommon.CWBCustomers
        Get

            If _CWBCustomers Is Nothing Then
                _CWBCustomers = New CWBCommon.CWBCustomers
            End If

            Return _CWBCustomers
        End Get
    End Property
    Public ReadOnly Property CWBSales() As CWBCommon.CWBSales
        Get

            If _CWBSales Is Nothing Then
                _CWBSales = New CWBCommon.CWBSales
            End If

            Return _CWBSales
        End Get
    End Property
#End Region

#Region "Form Events"
    Private Sub frmCustomerSummary_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.Chr(27) Then
            Me.Close()
        End If
    End Sub

    Private Sub frmCustomerSummary_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.deFromDate.EditValue = Date.Today
        Me.deToDate.EditValue = Date.Today
        Me.PopulateCustomerLookup()
    End Sub
#End Region

#Region "Populate Customer Lookup"
    Public Sub PopulateCustomerLookup()

        Try
            With leCustomers
                .Properties.DataSource = CWBCustomers.GetAllCustomers().Tables(0)
                .Properties.DisplayMember = "CustomerName"
                .Properties.ValueMember = "CustomerID"
                .Properties.PopupWidth = 350
            End With


        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Edit Button Events"
    Private Sub leCustomers_ButtonClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraEditors.Controls.ButtonPressedEventArgs) Handles leCustomers.ButtonClick
        Select Case e.Button.Index
            Case 1
                leCustomers.EditValue = Nothing
        End Select
    End Sub
#End Region

#Region "Buttton Events"

    Private Sub sbProcess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sbProcess.Click

        Try



            If leCustomers.Text <> String.Empty And teVehicleNo.Text <> String.Empty Then
                CWBSales.FromDate = Me.deFromDate.EditValue
                CWBSales.ToDate = Me.deToDate.EditValue

                CWBSales.CustomerID = Me.leCustomers.EditValue
                CWBSales.ReferenceNo = Me.teVehicleNo.Text.Trim
                gcCustomerSummary.DataSource = CWBSales.GetSalesByReferenceNoCustomerAndDates.Tables(0)

            End If


            If leCustomers.Text <> String.Empty And teVehicleNo.Text = String.Empty Then
                CWBSales.FromDate = Me.deFromDate.EditValue
                CWBSales.ToDate = Me.deToDate.EditValue

                CWBSales.CustomerID = Me.leCustomers.EditValue
                CWBSales.ReferenceNo = Me.teVehicleNo.Text.Trim
                gcCustomerSummary.DataSource = CWBSales.GetSalesByReferenceNoCustomerAndDates.Tables(1)

            End If


            If leCustomers.Text = String.Empty And teVehicleNo.Text <> String.Empty Then
                CWBSales.FromDate = Me.deFromDate.EditValue
                CWBSales.ToDate = Me.deToDate.EditValue

                CWBSales.CustomerID = Me.leCustomers.EditValue
                CWBSales.ReferenceNo = Me.teVehicleNo.Text.Trim
                gcCustomerSummary.DataSource = CWBSales.GetSalesByReferenceNoCustomerAndDates.Tables(2)

            End If


        Catch ex As Exception
            MessageError(ex.ToString)
        End Try

    End Sub

    Private Sub sbPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sbPrint.Click
        PrintPreview(gcCustomerSummary, "Customer Report")
    End Sub

    Private Sub sbSaveDiscount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sbSaveDiscount.Click
        Try

            If gvCustomerSummary.RowCount > 0 Then

                gvCustomerSummary.PostEditor()
                gvCustomerSummary.MoveLast()

                For i As Integer = 0 To Me.gvCustomerSummary.RowCount - 1

                    CWBSales.SalesID = Convert.ToInt64(Me.gvCustomerSummary.GetRowCellDisplayText(i, GridColumn9))
                    CWBSales.PDiscount = Convert.ToDecimal(Me.gvCustomerSummary.GetRowCellDisplayText(i, GridColumn5))
                    CWBSales.Discount = Convert.ToDecimal(Me.gvCustomerSummary.GetRowCellDisplayText(i, GridColumn6))
                    CWBSales.GrandTotal = Convert.ToDecimal(Me.gvCustomerSummary.GetRowCellDisplayText(i, GridColumn7))
                    CWBSales.UpdatedBy = UserID
                    CWBSales.SalesUpdateDiscount()

                Next

                Dim frm As New frmSavedOk
                frm.Text = CWB_SAVESUCCESS_CONFIRMATION_TITLE
                frm.lblTitle.Text = CWB_SAVESUCCESS_CONFIRMATION_TITLELABEL
                frm.lblDescription.Text = CWB_SAVESUCCESS_CONFIRMATION_DESCRIPTIONLABEL
                frm.ShowDialog()
            End If

        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Grid Events"
    Private Sub gvCustomerSummary_CellValueChanged(ByVal sender As System.Object, ByVal e As DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs) Handles gvCustomerSummary.CellValueChanged
        Try
            Select Case e.Column.VisibleIndex

                Case 5
                    Dim a, b, c As Decimal

                    a = IIf(Not IsDBNull(gvCustomerSummary.GetFocusedRowCellValue(GridColumn4)), gvCustomerSummary.GetFocusedRowCellValue(GridColumn4), 0)
                    b = IIf(Not IsDBNull(gvCustomerSummary.GetFocusedRowCellValue(GridColumn5)), gvCustomerSummary.GetFocusedRowCellValue(GridColumn5), 0)

                    gvCustomerSummary.SetFocusedRowCellValue(GridColumn6, a * b / 100)
                    gvCustomerSummary.SetFocusedRowCellValue(GridColumn7, ((a - (a * b / 100))))
                    a = 0
                    b = 0
                    c = 0
                    gvCustomerSummary.UpdateTotalSummary()
            End Select
        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
#End Region

    
End Class