' ============================================================================
'  Modules\DeviceActions.vb
'
'  Everything the Device menu and the device half of the Tools menu do.
'
'  Add Device and Remove Device are still placeholders - they report to the status
'  bar and nothing more. The three Device Editor commands are implemented.
' ============================================================================

Public Module DeviceActions

    ''' <summary>
    ''' Device / Add Device - pick a .DEV file and put that device on screen, below
    ''' the ones already there.
    ''' </summary>
    Public Sub AddDevice()

        Dim devicefile As String = DeviceLibrary.PickDeviceFile(AppCore.MainShell, "Add Device")

        If devicefile.Length = 0 Then
            AppCore.SetStatus("Add Device cancelled")
            Exit Sub
        End If

        AppCore.AddDevicePanel(devicefile)

    End Sub

    ''' <summary>Device / Remove Device - take the selected device out of the project.</summary>
    Public Sub RemoveDevice()

        ' TODO: remove the selected device, then AppCore.MarkModified().
        AppCore.ReportNotImplemented("Remove Device")

    End Sub

    ''' <summary>Tools / Create Device - a blank Device Editor.</summary>
    Public Sub CreateDevice()

        ShowDeviceEditor(String.Empty, DeviceEditorMode.mode_New)

    End Sub

    ''' <summary>
    ''' Tools / Edit Device - pick a .DEV file and edit it in place. The device name
    ''' is locked, because it is what names the file.
    ''' </summary>
    Public Sub EditDevice()

        Dim devicefile As String = DeviceLibrary.PickDeviceFile(AppCore.MainShell, "Edit Device")

        If devicefile.Length = 0 Then
            AppCore.SetStatus("Edit Device cancelled")
            Exit Sub
        End If

        ShowDeviceEditor(devicefile, DeviceEditorMode.mode_Edit)

    End Sub

    ''' <summary>
    ''' Tools / Clone Device - pick a .DEV file and open it as the starting point for
    ''' a new one. The name is flagged until it is changed, and saving over the
    ''' original is refused.
    ''' </summary>
    Public Sub CloneDevice()

        Dim devicefile As String = DeviceLibrary.PickDeviceFile(AppCore.MainShell, "Clone Device")

        If devicefile.Length = 0 Then
            AppCore.SetStatus("Clone Device cancelled")
            Exit Sub
        End If

        ShowDeviceEditor(devicefile, DeviceEditorMode.mode_Clone)

    End Sub

    ''' <summary>
    ''' Opens the editor modally. DeviceEditorCore reports its own success to the
    ''' status bar, so only the abandoned case is worth a message here.
    ''' </summary>
    Private Sub ShowDeviceEditor(devicefile As String, mode As DeviceEditorMode)

        Dim editor As New DeviceEditorWindow(devicefile, mode)
        editor.Owner = AppCore.MainShell

        Dim saved As Boolean? = editor.ShowDialog()

        If Not saved.GetValueOrDefault() Then AppCore.SetStatus("Device Editor closed without saving")

    End Sub

End Module
