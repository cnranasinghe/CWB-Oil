Imports CWBCommon.CWBLogBook
Public Class frmAudit

#Region "Variables"
    Dim _CWBSAudit As CWBCommon.CWBAudit

#End Region

#Region "Constructor"
    Public ReadOnly Property CWBAudit() As CWBCommon.CWBAudit
        Get
            ' Create on demand...
            If _CWBSAudit Is Nothing Then
                _CWBSAudit = New CWBCommon.CWBAudit()
            End If

            Return _CWBSAudit
        End Get
    End Property
#End Region

#Region "Form Events"
    Private Sub frmAudit_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.Chr(27) Then
            Me.Close()

        End If
    End Sub
    Private Sub frmAudit_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Me.deFromDate.EditValue = Date.Today
        'Me.deToDate.EditValue = Date.Today

        Try

            With CWBAudit

                Dim salesno As Int64
                salesno = Convert.ToInt64(lblSalesId.Text.Replace("System No - ", ""))

                gcLogBook.DataSource = .GetSalesAuditBySalesID(salesno).Tables(0)

            End With

        Catch ex As Exception
            MessageError(ex.ToString)
        End Try



    End Sub

#End Region

#Region "Button Events"
    Private Sub sbProcess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


        'Try
        '    If dxvpLogBook.Validate Then
        '        With CWBAudit
        '            .FromDate = Me.deFromDate.EditValue
        '            .ToDate = Me.deToDate.EditValue
        '            gcLogBook.DataSource = .GetLogBookByDates().Tables(0)
        '        End With

        '    End If

        'Catch ex As Exception
        '    MessageError(ex.ToString)
        'End Try

    End Sub
#End Region

#Region "Button Events"
    Private Sub sbPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sbPrint.Click
        Me.Close()
        PrintPreview(gcLogBook, "Audit Report for " + lblSalesId.Text)
    End Sub
#End Region
End Class