Imports CWBCommon.CWBCollections
Imports CWBCommon.CWBSaless
Imports Microsoft.Practices.EnterpriseLibrary.Common
Imports Microsoft.Practices.EnterpriseLibrary.Data
Imports System.Data
Imports System.Data.Common
Imports CWBCommon.CWBConstants
Imports CWBCommon.CWBModels
Public Class frmBulkSettlement


#Region "Variables"
    Dim _CWBCustomers As CWBCommon.CWBCustomers
    Dim _CWBSales As CWBCommon.CWBSales
    Dim _CWBCollections As CWBCommon.CWBCollections
#End Region

#Region "Constructor"
    Public ReadOnly Property CWBCustomers() As CWBCommon.CWBCustomers
        Get

            If _CWBCustomers Is Nothing Then
                _CWBCustomers = New CWBCommon.CWBCustomers
            End If

            Return _CWBCustomers
        End Get
    End Property
    Public ReadOnly Property CWBCollections() As CWBCommon.CWBCollections
        Get

            If _CWBCollections Is Nothing Then
                _CWBCollections = New CWBCommon.CWBCollections
            End If

            Return _CWBCollections
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
    Private Sub frmBulkSettlement_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.Chr(27) Then
            Me.Close()
        End If
    End Sub

    Private Sub frmBulkSettlement_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        deFromDate.EditValue = Date.Today
        deToDate.EditValue = Date.Today

    End Sub
    Private Sub frmBulkSettlement_Activated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        Me.PopulateCustomerLookup()
    End Sub
#End Region

#Region "Populate Customer Lookup"
    Public Sub PopulateCustomerLookup()

        Try
            With LeCustomers
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

#Region "Populate Unsettle Sales Bills"
    Private Sub PopulateUnsettleGrid()
        Try
            CWBSales.FromDate = Me.deFromDate.EditValue
            CWBSales.ToDate = Me.deToDate.EditValue
            CWBSales.CustomerID = Me.LeCustomers.EditValue
            gcUnsettledBills.DataSource = CWBSales.GetAllSalesUnsettledBillsByCustomerIDAndDates().Tables(0)

        Catch ex As Exception
            MessageError(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Button Events"
    Private Sub SimpleButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SimpleButton1.Click
        If dxvpBulkReceits.Validate Then
            Me.PopulateUnsettleGrid()
        End If
    End Sub

    Private Sub sbPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sbPrint.Click
        PrintPreview(gcUnsettledBills, "Customer Report")
    End Sub

    Private Sub sbSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sbSave.Click



        Dim frm As New frmUpdateYesNo
        frm.peImage.Image = CWB.My.Resources.Resources.ImgUpdate
        frm.Text = CWB_UPDATE_CONFIRMATION_TITLE
        frm.lblTitle.Text = CWB_UPDATE_CONFIRMATION_TITLELABEL
        frm.lblDescription.Text = CWB_UPDATE_CONFIRMATION_DESCRIPTIONLABEL



        If frm.ShowDialog = Windows.Forms.DialogResult.Yes Then
            Dim _Connection As DbConnection = Nothing
            Dim _Transaction As DbTransaction = Nothing
            Dim _DB As Database = DatabaseFactory.CreateDatabase(CWB_DBCONNECTION_STRING)
            _Connection = _DB.CreateConnection
            _Connection.Open()
            _Transaction = _Connection.BeginTransaction()

            Try
                With CWBCollections
                    For i As Integer = 0 To Me.gvUnsettledBills.RowCount - 1
                        If gvUnsettledBills.GetRowCellDisplayText(i, GridColumn7) > 0 Then

                            .SystemID = Me.gvUnsettledBills.GetRowCellDisplayText(i, GridColumn1)

                            .TransactionTypeID = CWBCommon.CWBEnums.EnumTransactionTypes.SALES
                            .PaymentTypeID = CWBCommon.CWBEnums.EnumPaymentTypes.CASH
                            .Date = Date.Today
                            .Reference = ""

                            .Amount = FormatNumber(Me.gvUnsettledBills.GetRowCellDisplayText(i, GridColumn7), 2, , , )
                            .InsertCollection(_DB, _Transaction)

                            CWBSales.SalesID = Me.gvUnsettledBills.GetRowCellDisplayText(i, GridColumn1)
                            CWBSales.UpdatedBy = UserID
                            CWBSales.SalesSetAsPaid()
                        ElseIf gvUnsettledBills.GetRowCellDisplayText(i, GridColumn7) = 0 Then

                            CWBSales.SalesID = Me.gvUnsettledBills.GetRowCellDisplayText(i, GridColumn1)
                            CWBSales.UpdatedBy = UserID
                            CWBSales.SalesSetAsPaid()
                        End If

                    Next



                End With

                _Transaction.Commit()

                Dim frmx As New frmSavedOk
                frmx.Text = CWB_SAVESUCCESS_CONFIRMATION_TITLE
                frmx.lblTitle.Text = CWB_SAVESUCCESS_CONFIRMATION_TITLELABEL
                frmx.lblDescription.Text = CWB_SAVESUCCESS_CONFIRMATION_DESCRIPTIONLABEL
                frmx.ShowDialog()
                Me.PopulateUnsettleGrid()

            Catch ex As Exception
                _Transaction.Rollback()
                MessageError(ex.ToString)
            End Try

        End If





    End Sub
#End Region




    
End Class