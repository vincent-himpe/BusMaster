' ============================================================================
'  Modules\WorkspaceCore.vb
'
'  The Workspace block on the toolbar: a list of saved views, a Save and a
'  Delete.
'
'  A view is a record of which registers had their padlock shut. Since a shut
'  padlock is what keeps a register on screen when its panel is collapsed, saving
'  one is really saving "show me these registers and nothing else" - a way to
'  keep several ways of looking at the same bus and switch between them.
'
'  Views live in "<project>.Workspace", beside the project file and separate from
'  it, so they can be added, thrown away or handed on without touching the layout
'  itself. There is nowhere to put them until the project has been saved once.
'
'  Panels are matched by identity, never by position or name:
'
'      in the view, not in the project   the panel has been deleted - skip it
'      in the project, not in the view   the panel is new - leave it alone
'
'  Neither is an error, which is what lets a view outlive the layout changing.
' ============================================================================

Imports System.Globalization
Imports System.IO
Imports Microsoft.VisualBasic

Public Module WorkspaceCore

    Public Const WorkspaceExtension As String = ".Workspace"

    Private Const DialogTitle As String = "Workspace"

    ''' <summary>The views for the open project, as last read or written.</summary>
    Private Views As WorkspaceFile

    ''' <summary>Guards the list while it is being refilled, so it does not apply itself.</summary>
    Private Refilling As Boolean = False


    ' ========================================================================
    '  Where the views live
    ' ========================================================================

    ''' <summary>
    ''' "<project>.Workspace", beside the project. An empty string while the project
    ''' has never been saved, because until then it has no name to hang views on.
    ''' </summary>
    Public ReadOnly Property WorkspacePath As String
        Get
            Dim projectPath As String = AppCore.CurrentProjectPath

            If String.IsNullOrWhiteSpace(projectPath) Then Return String.Empty

            Return Path.Combine(Path.GetDirectoryName(projectPath),
                                Path.GetFileNameWithoutExtension(projectPath) & WorkspaceExtension)
        End Get
    End Property


    ' ========================================================================
    '  Reading and writing
    ' ========================================================================

    ''' <summary>
    ''' Reads the views for whatever project is open and refills the toolbar list.
    ''' Called whenever the open project changes. A project with no views file yet
    ''' simply gets an empty list - that is the normal state, not a failure.
    ''' </summary>
    Public Sub Reload()

        Views = New WorkspaceFile

        Dim filePath As String = WorkspacePath

        If filePath.Length > 0 AndAlso File.Exists(filePath) Then
            Try
                Views = JsonFile.ReadFrom(Of WorkspaceFile)(filePath)
            Catch ex As Exception
                AppCore.SetStatus("Views not read (" & ex.Message & ")")
                Views = New WorkspaceFile
            End Try
        End If

        If Views.Views Is Nothing Then Views.Views = New List(Of WorkspaceView)

        RefreshList()

    End Sub

    Private Function Write() As Boolean

        Dim filePath As String = WorkspacePath
        If filePath.Length = 0 Then Return False

        Try
            JsonFile.WriteTo(filePath, Views)
            Return True

        Catch ex As Exception
            MessageBox.Show(AppCore.MainShell,
                            "The views could not be saved:" & vbCrLf & vbCrLf & ex.Message,
                            DialogTitle, MessageBoxButton.OK, MessageBoxImage.Error)
            Return False
        End Try

    End Function


    ' ========================================================================
    '  The toolbar list
    ' ========================================================================

    ''' <summary>Refills the drop-down from Views, leaving nothing selected.</summary>
    Private Sub RefreshList()

        Dim shell As MainWindow = AppCore.MainShell
        If shell Is Nothing Then Exit Sub

        Refilling = True

        shell.tlbr_Workspace_List.Items.Clear()

        For Each view As WorkspaceView In Views.Views
            If view IsNot Nothing Then shell.tlbr_Workspace_List.Items.Add(view.Name)
        Next

        shell.tlbr_Workspace_List.SelectedIndex = -1

        Refilling = False

    End Sub

    ''' <summary>Puts a name in the list and selects it, without applying it again.</summary>
    Private Sub SelectName(name As String)

        Dim shell As MainWindow = AppCore.MainShell
        If shell Is Nothing Then Exit Sub

        Refilling = True
        shell.tlbr_Workspace_List.SelectedItem = name
        Refilling = False

    End Sub

    ''' <summary>The name showing in the drop-down, or an empty string.</summary>
    Private Function SelectedName() As String

        Dim shell As MainWindow = AppCore.MainShell
        If shell Is Nothing Then Return String.Empty

        Return If(TryCast(shell.tlbr_Workspace_List.SelectedItem, String), String.Empty)

    End Function

    ''' <summary>Picking a view from the list puts it on screen.</summary>
    Public Sub ViewChosen()

        If Refilling Then Exit Sub

        Dim name As String = SelectedName()
        If name.Length = 0 Then Exit Sub

        ApplyView(name)

    End Sub


    ' ========================================================================
    '  Save
    ' ========================================================================

    ''' <summary>
    ''' Toolbar / Workspace / Save. Asks for a name, records every panel's padlocks
    ''' under it, and adds it to the list.
    ''' </summary>
    Public Sub SaveView()

        If Not HaveSomewhereToSave() Then Exit Sub

        Dim typed As String = StringPrompt.Ask(
            AppCore.MainShell, DialogTitle,
            "Name for this view:" & vbCrLf & vbCrLf &
            "Saves the visibility state for the registers.",
            SelectedName())

        If typed Is Nothing Then
            AppCore.SetStatus("View not saved")
            Exit Sub
        End If

        Dim entered As String = typed.Trim()

        If entered.Length = 0 Then
            AppCore.SetStatus("View not saved")
            Exit Sub
        End If

        Dim existing As WorkspaceView = Find(entered)

        If existing IsNot Nothing Then

            Dim answer As MessageBoxResult = MessageBox.Show(
                AppCore.MainShell,
                "There is already a view called """ & entered & """." & vbCrLf & vbCrLf & "Replace it?",
                DialogTitle, MessageBoxButton.OKCancel, MessageBoxImage.Question)

            If answer <> MessageBoxResult.OK Then
                AppCore.SetStatus("View not saved")
                Exit Sub
            End If

            existing.Panels = CollectPanels()

        Else

            Views.Views.Add(New WorkspaceView With {
                .Name = entered,
                .Panels = CollectPanels()
            })

        End If

        If Not Write() Then Exit Sub

        RefreshList()
        SelectName(entered)

        AppCore.SetStatus("Saved view " & entered)

    End Sub

    ''' <summary>
    ''' Toolbar / Workspace / Save, right-clicked. Writes what is on screen over the
    ''' view showing in the list.
    '''
    ''' No name to type and nothing to confirm - that is the whole point of it. The
    ''' left click is the one that makes something new and therefore has to ask.
    ''' </summary>
    Public Sub UpdateView()

        If Not HaveSomewhereToSave() Then Exit Sub

        Dim name As String = SelectedName()

        If name.Length = 0 Then
            AppCore.SetStatus("Pick a view to update, or left-click Save to make a new one")
            Exit Sub
        End If

        Dim view As WorkspaceView = Find(name)
        If view Is Nothing Then Exit Sub

        view.Panels = CollectPanels()

        If Not Write() Then Exit Sub

        AppCore.SetStatus("Updated view " & name)

    End Sub

    ''' <summary>Every panel on screen and the registers it has locked.</summary>
    Private Function CollectPanels() As List(Of WorkspacePanelState)

        Dim states As New List(Of WorkspacePanelState)

        Dim shell As MainWindow = AppCore.MainShell
        If shell Is Nothing Then Return states

        For Each child As UIElement In shell.stk_Devices.Children

            Dim panel As DevicePanel = TryCast(child, DevicePanel)
            If panel Is Nothing Then Continue For

            states.Add(New WorkspacePanelState With {
                .PanelId = If(panel.PanelId, String.Empty),
                .LockedRegisters = panel.GetLockedRegisters()
            })

        Next

        Return states

    End Function


    ' ========================================================================
    '  Restore
    ' ========================================================================

    ''' <summary>
    ''' Puts a saved view back on screen. Panels the view does not mention keep
    ''' whatever they have set now - they were added after it was saved, and there
    ''' is nothing recorded to put back.
    ''' </summary>
    Public Sub ApplyView(name As String)

        Dim view As WorkspaceView = Find(name)
        If view Is Nothing Then Exit Sub

        Dim shell As MainWindow = AppCore.MainShell
        If shell Is Nothing Then Exit Sub

        Dim restored As Integer = 0
        Dim missing As Integer = 0

        For Each state As WorkspacePanelState In view.Panels

            If state Is Nothing Then Continue For

            Dim panel As DevicePanel = PanelWithId(state.PanelId)

            ' Named by the view but no longer on screen: the user has removed it.
            If panel Is Nothing Then
                missing += 1
                Continue For
            End If

            panel.SetLockedRegisters(state.LockedRegisters)
            restored += 1

        Next

        Dim report As String = "View " & name & " - " &
                               restored.ToString(CultureInfo.InvariantCulture) & " devices restored"

        If missing > 0 Then
            report &= ", " & missing.ToString(CultureInfo.InvariantCulture) & " no longer here"
        End If

        AppCore.SetStatus(report)

    End Sub

    Private Function PanelWithId(panelId As String) As DevicePanel

        If String.IsNullOrWhiteSpace(panelId) Then Return Nothing

        Dim shell As MainWindow = AppCore.MainShell
        If shell Is Nothing Then Return Nothing

        For Each child As UIElement In shell.stk_Devices.Children

            Dim panel As DevicePanel = TryCast(child, DevicePanel)
            If panel Is Nothing Then Continue For

            If String.Equals(panel.PanelId, panelId, StringComparison.OrdinalIgnoreCase) Then Return panel

        Next

        Return Nothing

    End Function


    ' ========================================================================
    '  Delete
    ' ========================================================================

    ''' <summary>
    ''' Toolbar / Workspace / Delete. Drops the view showing in the list, and its
    ''' record with it. Nothing on screen changes - a view is a bookmark, not the
    ''' thing itself.
    ''' </summary>
    Public Sub DeleteView()

        Dim name As String = SelectedName()

        If name.Length = 0 Then
            AppCore.SetStatus("Pick a view to delete first")
            Exit Sub
        End If

        Dim view As WorkspaceView = Find(name)
        If view Is Nothing Then Exit Sub

        Views.Views.Remove(view)

        If Not Write() Then Exit Sub

        RefreshList()

        AppCore.SetStatus("Deleted view " & name)

    End Sub


    ' ========================================================================
    '  Shared helpers
    ' ========================================================================

    Private Function Find(name As String) As WorkspaceView

        If Views Is Nothing OrElse Views.Views Is Nothing Then Return Nothing

        For Each view As WorkspaceView In Views.Views
            If view Is Nothing Then Continue For
            If String.Equals(view.Name, name, StringComparison.OrdinalIgnoreCase) Then Return view
        Next

        Return Nothing

    End Function

    ''' <summary>
    ''' Views are named after the project file, so there has to be one. Says so
    ''' rather than writing them somewhere the user will not find them again.
    ''' </summary>
    Private Function HaveSomewhereToSave() As Boolean

        If Views Is Nothing Then Views = New WorkspaceFile

        If WorkspacePath.Length > 0 Then Return True

        MessageBox.Show(AppCore.MainShell,
                        "Views are kept beside the project file, so the project has to be " &
                        "saved before one can be made.",
                        DialogTitle, MessageBoxButton.OK, MessageBoxImage.Information)

        AppCore.SetStatus("Save the project before saving a view")

        Return False

    End Function

End Module
