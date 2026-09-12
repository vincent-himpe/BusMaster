' ============================================================================
'  Modules\ProbeFirmware.vb
'
'  Tools / Probe Firmware - putting firmware on the probe itself.
'
'  The probe is an STM32, and it is programmed over DFU: held in the bootloader
'  by the BOOT button while RESET is pulsed, it comes up on USB as a DFU device
'  rather than as the serial port the rest of the program talks to. That is a
'  different conversation from anything in ProbeControl, which is why none of this
'  goes near it - provisioning happens with the probe *not* being a probe.
'
'  Two entries, and they are not the same job:
'
'      Provision Probe   a blank processor, straight from the reel. The user has
'                        to put it in the bootloader by hand, so the window says
'                        how, and everything is done through the DFU device.
'      Update Firmware   a probe that is already running and can be asked to
'                        reboot into the bootloader itself.
'
'  **The work is not written yet.** Flashing means driving ST's own command line
'  programmer, and the user is providing the sample code that finds devices with
'  it. Until then Scan Devices and Flash Firmware say so in the window rather than
'  doing something approximate - a firmware tool that half works is worse than one
'  that admits it does not, because the thing on the other end is a part that can
'  be bricked.
'
'  Everything the window can do lives here; ProvisionProbeWindow.xaml.vb is one
'  line per button.
' ============================================================================

Imports System.IO
Imports Microsoft.Win32

Public Module ProbeFirmware

    Private Const DialogTitle As String = "Provision Probe"

    ''' <summary>What the file browser will open, and what a firmware image is.</summary>
    Private Const FirmwareExtension As String = ".bin"
    Private Const FirmwareFilter As String = "Firmware image (*.bin)|*.bin|All Files (*.*)|*.*"

    Private Const NoFirmware As String = "No firmware selected"


    ''' <summary>Theme keys for the three things the status line has to say.</summary>
    Private Const MissingInk As String = "Brush_Provision_Missing"
    Private Const FoundInk As String = "Brush_Provision_Found"
    Private Const DoneInk As String = "Brush_Provision_Done"


    Private Provision As ProvisionProbeWindow

    ''' <summary>The firmware image the user picked, or an empty string.</summary>
    Private FirmwarePath As String = String.Empty

    ''' <summary>
    ''' The serial number the last scan found, or an empty string for "we do not
    ''' know of a part to flash".
    '''
    ''' Cleared after a flash as well as after a failed scan, because the part
    ''' reboots out of the bootloader once it has been written to - the serial is
    ''' then a number that used to be true. Scanning again is what proves there is
    ''' still something there, and until that has been done there is nothing to
    ''' flash.
    ''' </summary>
    Private DfuSerial As String = String.Empty

    ''' <summary>
    ''' Where the file browser opens next time. Remembered for as long as the
    ''' program runs but deliberately not saved: firmware is picked up from
    ''' wherever it was built, which is not somewhere to make a setting of.
    ''' </summary>
    Private LastFirmwareFolder As String = String.Empty


    ' ========================================================================
    '  Opening the window
    ' ========================================================================

    ''' <summary>
    ''' Tools / Probe Firmware / Provision Probe. Modal, like the program's other
    ''' dialogs - nothing else can be done to a probe that is sitting in its
    ''' bootloader, so there is nothing to do behind this window anyway.
    ''' </summary>
    Public Sub ProvisionProbe()

        Dim dialog As New ProvisionProbeWindow With {.Owner = AppCore.MainShell}

        ModalShade.Cover(dialog)

        Try
            dialog.ShowDialog()
        Finally
            ModalShade.Uncover()
            Provision = Nothing
        End Try

    End Sub

    ''' <summary>
    ''' Tools / Probe Firmware / Update Firmware - reflash a probe that is already
    ''' running, which can be asked to enter the bootloader rather than having its
    ''' buttons held down.
    ''' </summary>
    Public Sub UpdateFirmware()

        ' TODO: ask the running probe to reboot into DFU, then flash it the same way
        ' Provision does. Waiting on the same programmer code as the scan.
        AppCore.ReportNotImplemented("Update Firmware")

    End Sub

    ''' <summary>
    ''' Called from the window's Loaded handler. Anything that has to be true the
    ''' moment it appears is set here rather than in the XAML, so there is one place
    ''' that decides what state the window opens in.
    ''' </summary>
    Public Sub Attach(window As ProvisionProbeWindow)

        Provision = window

        WindowTheme.ApplyDarkTitleBar(Provision)

        ' A fresh window asks for the firmware again. Which image belongs on a part
        ' is the whole question here, and carrying the last answer into a new job
        ' would answer it for the user without their noticing.
        FirmwarePath = String.Empty

        ' And a part found last time is not a part that is there now.
        DfuSerial = String.Empty

        ShowFirmware()
        ClearStatus()

        AddHandler Provision.PreviewKeyDown, AddressOf Window_PreviewKeyDown

    End Sub

    ''' <summary>
    ''' Escape closes the window, the same as the Close button. Taken here rather
    ''' than by marking Close IsCancel, because that property closes the window
    ''' itself on top of whatever the Click handler did.
    ''' </summary>
    Private Sub Window_PreviewKeyDown(sender As Object, e As KeyEventArgs)

        If e.Key <> Key.Escape Then Exit Sub

        CloseProvision()
        e.Handled = True

    End Sub

    Public Sub CloseProvision()

        If Provision Is Nothing Then Exit Sub

        Provision.Close()

    End Sub


    ' ========================================================================
    '  Choosing the image
    ' ========================================================================

    ''' <summary>
    ''' Picks the .BIN to flash. Nothing is read from it here - the programmer
    ''' takes the path - so a file that is not really firmware is not found out
    ''' until the flash is attempted.
    ''' </summary>
    Public Sub SelectFirmware()

        If Provision Is Nothing Then Exit Sub

        Dim dialog As New OpenFileDialog With {
            .Title = "Select Firmware",
            .Filter = FirmwareFilter,
            .DefaultExt = FirmwareExtension,
            .CheckFileExists = True
        }

        If LastFirmwareFolder.Length > 0 AndAlso Directory.Exists(LastFirmwareFolder) Then
            dialog.InitialDirectory = LastFirmwareFolder
        End If

        Dim picked As Boolean? = dialog.ShowDialog(Provision)

        If Not picked.GetValueOrDefault() Then Exit Sub

        FirmwarePath = dialog.FileName

        Try
            LastFirmwareFolder = Path.GetDirectoryName(FirmwarePath)
        Catch
            LastFirmwareFolder = String.Empty
        End Try

        ShowFirmware()

        ' Whatever the status line was saying was about the last file or the last
        ' flash, and neither is this one. The scan result goes with it, which is a
        ' small loss against leaving "upload success" sitting over a different file.
        ClearStatus()

    End Sub

    ''' <summary>Puts the chosen path under the button.</summary>
    Private Sub ShowFirmware()

        If Provision Is Nothing Then Exit Sub

        Provision.lbl_FirmwarePath.Text = If(FirmwarePath.Length > 0, FirmwarePath, NoFirmware)

        RefreshFlashButton()

    End Sub

    ''' <summary>
    ''' Flashing needs both halves: an image to write, and a part known to be
    ''' sitting in the bootloader waiting for it. Either one on its own is not
    ''' enough, so this is the only place the button is switched and every routine
    ''' that changes either half calls it.
    ''' </summary>
    Private Sub RefreshFlashButton()

        If Provision Is Nothing Then Exit Sub

        Provision.btn_FlashFirmware.IsEnabled = FirmwarePath.Length > 0 AndAlso DfuSerial.Length > 0

    End Sub


    ' ========================================================================
    '  The programmer
    '
    '  Both of these are one call into Modules\STM32DFU.vb and a reading of what
    '  came back. Nothing here knows how the part is actually reached - see that
    '  file for the bodies, which are still to be written.
    ' ========================================================================

    ''' <summary>
    ''' Looks for an STM32 sitting in DFU mode, and says what it found.
    '''
    ''' Finding nothing is the ordinary outcome, not an error: it is what the BOOT
    ''' and RESET dance not taking looks like, and the user's next move is to do it
    ''' again rather than to worry.
    ''' </summary>
    Public Sub ScanDevices()

        ' Nothing can be asked of a part without ST's programmer on the machine, so
        ' that is looked for first and said plainly if it is missing. Told apart
        ' from finding no device on purpose: one is fixed by installing something,
        ' the other by holding a button down, and a single "scan failed" would leave
        ' the user guessing which.
        If STM32DFU.FindProgrammer().Length = 0 Then

            DfuSerial = String.Empty

            ShowStatus("STM32Cube is not installed. This is required", MissingInk)
            RefreshFlashButton()

            Exit Sub

        End If

        DfuSerial = If(STM32DFU.GetDFUSerialNumber(), String.Empty).Trim()

        If DfuSerial.Length = 0 Then
            ShowStatus("No DFU Devices found", MissingInk)
        Else
            ShowStatus("Found DFU Serial " & DfuSerial, FoundInk)
        End If

        RefreshFlashButton()

    End Sub

    ''' <summary>
    ''' Writes the chosen image to the part the scan found. Only reachable with both
    ''' of those in hand - see RefreshFlashButton.
    '''
    ''' The button goes out again either way. On a success the part has rebooted and
    ''' is no longer in the bootloader; on a failure nobody knows what state it is
    ''' in. Both want the same thing next, which is a fresh scan, and leaving the
    ''' button live would invite a second flash at a serial number that has stopped
    ''' being true.
    ''' </summary>
    Public Sub FlashFirmware()

        If FirmwarePath.Length = 0 OrElse DfuSerial.Length = 0 Then Exit Sub

        Dim outcome As String = If(STM32DFU.FlashFile(FirmwarePath), String.Empty).Trim()

        ' Anything that is not a plain success is a failure. With a part that can be
        ' left unbootable, an answer nobody recognises is not worth the benefit of
        ' the doubt.
        If String.Equals(outcome, STM32DFU.Succeeded, StringComparison.OrdinalIgnoreCase) Then
            ShowStatus("Firmware upload success", DoneInk)
        Else
            ShowStatus("Flashing Failed", MissingInk)
        End If

        DfuSerial = String.Empty
        RefreshFlashButton()

    End Sub


    ' ========================================================================
    '  Small helpers
    ' ========================================================================

    ''' <summary>
    ''' The window's own status line. The main window's status bar is behind a modal
    ''' dialog and greyed out, so anything said there would not be read.
    '''
    ''' The colour carries as much as the words do - red for nothing there or it did
    ''' not work, green for found, blue for done - so it is given with the text
    ''' rather than left for the caller to remember separately.
    ''' </summary>
    Private Sub ShowStatus(text As String, brushKey As String)

        If Provision Is Nothing Then Exit Sub

        Provision.lbl_Status.Text = If(text, String.Empty)

        Dim ink As Brush = TryCast(Provision.TryFindResource(brushKey), Brush)
        If ink IsNot Nothing Then Provision.lbl_Status.Foreground = ink

    End Sub

    ''' <summary>
    ''' Empties the status line. The colour is left where it was - there is nothing
    ''' to read, so nothing to colour.
    ''' </summary>
    Private Sub ClearStatus()

        If Provision Is Nothing Then Exit Sub

        Provision.lbl_Status.Text = String.Empty

    End Sub

End Module
