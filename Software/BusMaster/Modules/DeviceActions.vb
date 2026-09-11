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
    ''' A device file has just been written, so every panel on screen showing that
    ''' file is built again from it. Returns how many were reloaded.
    '''
    ''' Called by the Device Editor on a successful save: without this, editing a
    ''' device leaves the panels already on screen showing the register list the
    ''' file used to have until the project is opened again.
    '''
    ''' Matching is on the resolved path, so the same device reached by a bare name
    ''' and by a full path is still the same device. Each panel keeps its own
    ''' function name, target, switches, values and open state - see
    ''' DevicePanel.ReloadDevice.
    ''' </summary>
    Public Function ReloadDeviceFile(devicefile As String) As Integer

        If String.IsNullOrWhiteSpace(devicefile) Then Return 0

        Dim wanted As String

        Try
            wanted = IO.Path.GetFullPath(devicefile)
        Catch
            Return 0
        End Try

        Dim reloaded As Integer = 0

        For Each panel As DevicePanel In AppCore.DevicePanels()

            If panel.DeviceFile.Length = 0 Then Continue For

            Dim showing As String

            Try
                showing = IO.Path.GetFullPath(panel.DeviceFile)
            Catch
                Continue For
            End Try

            If Not String.Equals(showing, wanted, StringComparison.OrdinalIgnoreCase) Then Continue For

            panel.ReloadDevice()
            reloaded += 1

        Next

        Return reloaded

    End Function


    ''' <summary>
    ''' Opens the editor modally. DeviceEditorCore reports its own success to the
    ''' status bar, so only the abandoned case is worth a message here.
    ''' </summary>
    Private Sub ShowDeviceEditor(devicefile As String, mode As DeviceEditorMode)

        Dim editor As New DeviceEditorWindow(devicefile, mode)
        editor.Owner = AppCore.MainShell

        ModalShade.Cover(editor)

        Dim saved As Boolean?

        Try
            saved = editor.ShowDialog()
        Finally
            ModalShade.Uncover()
        End Try

        If Not saved.GetValueOrDefault() Then AppCore.SetStatus("Device Editor closed without saving")

    End Sub

End Module
