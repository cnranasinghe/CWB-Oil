<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmError
    Inherits DevExpress.XtraEditors.XtraForm

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.lblError = New DevExpress.XtraEditors.LabelControl
        Me.sbOk = New DevExpress.XtraEditors.SimpleButton
        Me.meErrorMessage = New DevExpress.XtraEditors.MemoEdit
        CType(Me.meErrorMessage.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblError
        '
        Me.lblError.Appearance.BackColor = System.Drawing.Color.FromArgb(CType(CType(254, Byte), Integer), CType(CType(216, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.lblError.Appearance.BorderColor = System.Drawing.Color.Red
        Me.lblError.Appearance.Font = New System.Drawing.Font("Tahoma", 11.0!, System.Drawing.FontStyle.Bold)
        Me.lblError.Appearance.ForeColor = System.Drawing.Color.Red
        Me.lblError.Appearance.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblError.Appearance.Options.UseBackColor = True
        Me.lblError.Appearance.Options.UseBorderColor = True
        Me.lblError.Appearance.Options.UseFont = True
        Me.lblError.Appearance.Options.UseForeColor = True
        Me.lblError.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None
        Me.lblError.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Simple
        Me.lblError.LineColor = System.Drawing.Color.Red
        Me.lblError.Location = New System.Drawing.Point(73, 22)
        Me.lblError.Name = "lblError"
        Me.lblError.Padding = New System.Windows.Forms.Padding(5)
        Me.lblError.Size = New System.Drawing.Size(498, 47)
        Me.lblError.TabIndex = 1
        Me.lblError.Text = "Unexpected Error occured, Please contact System Administrator!"
        '
        'sbOk
        '
        Me.sbOk.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.sbOk.Location = New System.Drawing.Point(532, 276)
        Me.sbOk.Name = "sbOk"
        Me.sbOk.Size = New System.Drawing.Size(100, 30)
        Me.sbOk.TabIndex = 2
        Me.sbOk.Text = "Ok"
        '
        'meErrorMessage
        '
        Me.meErrorMessage.Location = New System.Drawing.Point(12, 84)
        Me.meErrorMessage.Name = "meErrorMessage"
        Me.meErrorMessage.Properties.ReadOnly = True
        Me.meErrorMessage.Size = New System.Drawing.Size(620, 172)
        Me.meErrorMessage.TabIndex = 3
        '
        'frmError
        '
        Me.AcceptButton = Me.sbOk
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.sbOk
        Me.ClientSize = New System.Drawing.Size(644, 318)
        Me.Controls.Add(Me.meErrorMessage)
        Me.Controls.Add(Me.sbOk)
        Me.Controls.Add(Me.lblError)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.KeyPreview = True
        Me.LookAndFeel.SkinName = "Black"
        Me.LookAndFeel.UseDefaultLookAndFeel = False
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmError"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "System Error, Can not proceed"
        Me.TopMost = True
        CType(Me.meErrorMessage.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lblError As DevExpress.XtraEditors.LabelControl
    Friend WithEvents sbOk As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents meErrorMessage As DevExpress.XtraEditors.MemoEdit
End Class
