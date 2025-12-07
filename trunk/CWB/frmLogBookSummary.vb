Imports CWBCommon.CWBLogBook
Public Class frmLogBookSummary

#Region "Variables"
    Dim _CWBLogBook As CWBCommon.CWBLogBook

#End Region

#Region "Constructor"
    Public ReadOnly Property CWBLogBook() As CWBCommon.CWBLogBook
        Get
            ' Create on demand...
            If _CWBLogBook Is Nothing Then
                _CWBLogBook = New CWBCommon.CWBLogBook()
            End If

            Return _CWBLogBook
        End Get
    End Property
#End Region

#Region "Form Events"
    Private Sub frmLogBookSummary_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.Chr(27) Then
            Me.Close()

        End If
    End Sub
    Private Sub frmLogBookSummary_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.deFromDate.EditValue = Date.Today
        Me.deToDate.EditValue = Date.Today
    End Sub

#End Region

#Region "Button Events"
    Private Sub sbProcess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sbProcess.Click


        Try
            If dxvpLogBook.Validate Then
                With CWBLogBook
                    .FromDate = Me.deFromDate.EditValue
                    .ToDate = Me.deToDate.EditValue
                    gcLogBook.DataSource = .GetLogBookByDates().Tables(0)
                End With

            End If

        Catch ex As Exception
            MessageError(ex.ToString)
        End Try

    End Sub
#End Region

#Region "Button Events"
    Private Sub sbPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sbPrint.Click
        PrintPreview(gcLogBook, "Login Report")
    End Sub
#End Region
End Class