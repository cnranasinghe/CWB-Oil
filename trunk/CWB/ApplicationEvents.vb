Imports System.Configuration
Namespace My

    ' The following events are available for MyApplication:
    ' 
    ' Startup: Raised when the application starts, before the startup form is created.
    ' Shutdown: Raised after all application forms are closed.  This event is not raised if the application terminates abnormally.
    ' UnhandledException: Raised if the application encounters an unhandled exception.
    ' StartupNextInstance: Raised when launching a single-instance application and the application is already active. 
    ' NetworkAvailabilityChanged: Raised when the network connection is connected or disconnected.
    Partial Friend Class MyApplication



#Region "Application Start up"
        Private Sub MyApplication_Startup(ByVal sender As Object, ByVal e As Microsoft.VisualBasic.ApplicationServices.StartupEventArgs) Handles Me.Startup
            Try


                If System.Configuration.ConfigurationManager.AppSettings("IsNetworkEnabled").ToString() = "True" Then
                    Dim isAvailable As Boolean
                    isAvailable = My.Computer.Network.IsAvailable

                    If isAvailable = False Then
                        frmNetworkAvailable.ShowDialog()
                    Else
                        Application.DoEvents()
                    End If
                End If

            Catch ex As ApplicationException
                MessageError(ex.ToString)
            End Try
        End Sub
#End Region

#Region "Raise When loss of Network"
        Private Sub MyApplication_NetworkAvailabilityChanged(ByVal sender As Object, ByVal e As Microsoft.VisualBasic.Devices.NetworkAvailableEventArgs) Handles Me.NetworkAvailabilityChanged

            If System.Configuration.ConfigurationManager.AppSettings("IsNetworkEnabled").ToString() = "True" Then
                Try
                    Dim isAvailable As Boolean
                    isAvailable = My.Computer.Network.IsAvailable

                    If isAvailable = False Then
                        frmNetworkAvailable.ShowDialog()
                    End If
                Catch ex As ApplicationException
                    MessageError(ex.ToString)
                End Try
            End If


        End Sub
#End Region

    End Class

End Namespace

