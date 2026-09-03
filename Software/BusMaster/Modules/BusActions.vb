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
    ''' Toolbar / Write All - push every register of every device on screen out to
    ''' the hardware in one go, whatever the Operation setting says about when
    ''' individual changes are written.
    ''' </summary>
    Public Sub WriteAll()

        ' TODO: walk the device panels and write each register map to the bus.
        AppCore.ReportNotImplemented("Write All")

    End Sub

End Module
