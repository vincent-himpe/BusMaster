' ============================================================================
'  ProvisionProbeWindow.xaml.vb
'
'  Event handlers ONLY - see Modules\ProbeFirmware.vb.
' ============================================================================

Class ProvisionProbeWindow

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        ProbeFirmware.Attach(Me)
    End Sub

    Private Sub btn_SelectFirmware_Click(sender As Object, e As RoutedEventArgs)
        ProbeFirmware.SelectFirmware()
    End Sub

    Private Sub btn_ScanDevices_Click(sender As Object, e As RoutedEventArgs)
        ProbeFirmware.ScanDevices()
    End Sub

    Private Sub btn_FlashFirmware_Click(sender As Object, e As RoutedEventArgs)
        ProbeFirmware.FlashFirmware()
    End Sub

    Private Sub btn_Close_Click(sender As Object, e As RoutedEventArgs)
        ProbeFirmware.CloseProvision()
    End Sub

End Class
