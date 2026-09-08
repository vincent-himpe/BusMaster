' ============================================================================
'  Modules\BusActions.vb
'
'  Everything the Bus menu does - the operations that talk to the I2C hardware.
'
'  The routines are wired up and reachable from the menu, but the bodies are
'  still placeholders - they report to the status bar and nothing more. Replace
'  one body at a time as the real work lands.
' ============================================================================

''' <summary>When a change made on screen is written out to the bus.</summary>
Public Enum BusOperationMode
    mode_RegisterOnChange = 0
    mode_DeviceOnChange = 1
End Enum

Public Module BusActions

    ''' <summary>Bus / Scan Bus - walk the address range and list what answers.</summary>
    Public Sub ScanBus()

        ' TODO: address sweep, then report how many devices answered.
        AppCore.ReportNotImplemented("Scan Bus")

    End Sub

    ''' <summary>Bus / Poll Bus - read the configured devices repeatedly.</summary>
    Public Sub PollBus()

        ' TODO: start and stop a repeating read of the project's devices.
        AppCore.ReportNotImplemented("Poll Bus")

    End Sub

    ''' <summary>Bus / Reset Bus - clear a hung bus and put the master back to idle.</summary>
    Public Sub ResetBus()

        ' TODO: clock out a stuck slave, then issue a stop condition.
        AppCore.ReportNotImplemented("Reset Bus")

    End Sub

    ''' <summary>
    ''' Toolbar / Write All - push every register of every device on screen out in
    ''' one go, whatever the Operation setting says about when individual changes
    ''' are written. That setting decides how much follows a change; this is asked
    ''' for outright.
    ''' </summary>
    Public Sub WriteAll()

        BusEvents.WriteSystem()

        AppCore.SetStatus("Wrote every register on " & Describe(AppCore.DevicePanels().Count))

    End Sub

    ''' <summary>
    ''' Toolbar / Read All - fetch every register of every device on screen. The
    ''' whole-workspace version of Ctrl+right-clicking one device.
    ''' </summary>
    Public Sub ReadAll()

        BusEvents.ReadSystem()

        AppCore.SetStatus("Read every register on " & Describe(AppCore.DevicePanels().Count))

    End Sub

    ''' <summary>"3 devices", or "1 device" - the status bar should read properly.</summary>
    Private Function Describe(count As Integer) As String

        If count = 1 Then Return "1 device"

        Return count.ToString(Globalization.CultureInfo.InvariantCulture) & " devices"

    End Function

End Module
