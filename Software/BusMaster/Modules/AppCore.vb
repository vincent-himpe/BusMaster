' ============================================================================
'  Modules\AppCore.vb
'
'  All of the application's active code. The window's event handlers do nothing
'  but call into here, which is why "Open" can be driven from the File menu, the
'  toolbar button and Ctrl+O without three copies of the logic.
'
'  Public entry points:
'
'      Attach              wire the module to the window (called once, on Loaded)
'      NewProject          start an empty project
'      OpenProject         pick a project file and load it
'      OpenRecentProject   load the project behind an MRU menu item
'      SaveProject         save to the current path, or ask if there is none
'      SaveProjectAs       always ask for a path
'      ShowHelp            show the help box
'      ExitApplication     close the window
'      MarkModified        flag the open project as having unsaved changes
'      SetStatus           write to the right-hand status bar field
' ============================================================================

Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
Imports Microsoft.Win32

Public Module AppCore

    Public Const ProjectExtension As String = ".Busmaster"
    Private Const ProjectFilter As String = "BusMaster Project (*.Busmaster)|*.Busmaster|All Files (*.*)|*.*"

    Public Const ProjectFolderName As String = "projects"
    Private Const UntitledName As String = "Untitled"
    Private Const AppName As String = "BusMaster"

    ''' <summary>
    ''' TEMPORARY. Loaded into the workspace panel at start-up so the layout can be
    ''' looked at. Delete this and the call in Attach once devices are opened for
    ''' real from a project.
    ''' </summary>
    Private Const DemoDeviceName As String = "PCA9555"

    ''' <summary>
    ''' The program's display name - "BusMaster 0.1". Used for the title bar and as
    ''' the caption of every message box.
    ''' </summary>
    Private ReadOnly Property AppTitle As String
        Get
            Return AppName & " " & AppSettings.ApplicationVersion
        End Get
    End Property

    ' The window this module drives. Set once by Attach.
    Private Shell As MainWindow

    ''' <summary>
    ''' The main window, for anything that needs an owner for a dialog.
    ''' </summary>
    Public ReadOnly Property MainShell As MainWindow
        Get
            Return Shell
        End Get
    End Property

    ''' <summary>The project currently open, or Nothing when there is none.</summary>
    Public Property CurrentProject As ProjectData

    ''' <summary>Where the open project lives, or an empty string if never saved.</summary>
    Public Property CurrentProjectPath As String = String.Empty

    ''' <summary>True when the open project has changes that are not on disk.</summary>
    Public Property IsModified As Boolean = False


    ' ========================================================================
    '  Start-up and shutdown
    ' ========================================================================

    ''' <summary>
    ''' Connects the module to the main window and brings the UI up to date.
    ''' Called from Window_Loaded.
    ''' </summary>
    Public Sub Attach(window As MainWindow)

        Shell = window

        WindowTheme.ApplyDarkTitleBar(Shell)

        If AppSettings.Load() Then
            SetStatus("Ready")
        Else
            SetStatus("Settings not read (" & AppSettings.LastError & ") - using defaults")
        End If

        RestoreWindowGeometry()
        ApplyWorkspaceBackground()
        RefreshRecentMenu()
        UpdateActiveProject()

        RefreshValueDisplayControls()
        RefreshProbePorts()

        ' Nothing is connected yet, so this is what greys the bus and target blocks
        ' out to start with.
        RefreshProbeConnection()

        ' One handler for every register on screen, now and later. Attached before
        ' any device is loaded so nothing can slip past it.
        BusEvents.Initialise()

        ' Built now but kept out of sight, so the log is already collecting whenever
        ' the user asks for it.
        CreateEventLogWindow()
        CreateCommandWindow()
        CreateTerminalWindow()

        ' Pick up where the user left off.
        If Not LoadMostRecentProject() Then

            ' TEMPORARY - see DemoDeviceName. A missing file just raises LoadFailed and
            ' writes to the status bar, so this cannot stop the program starting.
            Shell.dvp_Device.LoadDevice(DemoDeviceName)

            ' Opening a project does this itself; starting without one still needs
            ' the views list to exist and be empty.
            WorkspaceCore.Reload()

        End If

    End Sub

    ''' <summary>
    ''' Opens the project at the top of the recent list. Nothing here interrupts the
    ''' user: a list that is empty, or whose newest entry has been moved or deleted,
    ''' just leaves the program to start empty and says so in the status bar.
    ''' Returns True only if a project was actually opened.
    ''' </summary>
    Private Function LoadMostRecentProject() As Boolean

        Dim recent As List(Of String) = AppSettings.Current.RecentProjects

        If recent Is Nothing OrElse recent.Count = 0 Then Return False

        Dim projectPath As String = recent(0)
        If String.IsNullOrWhiteSpace(projectPath) Then Return False

        If Not File.Exists(projectPath) Then
            AppSettings.RemoveRecentProject(projectPath)
            RefreshRecentMenu()
            SetStatus("Last project not found - " & Path.GetFileName(projectPath))
            Return False
        End If

        LoadProjectFile(projectPath)

        Return CurrentProject IsNot Nothing

    End Function

    ''' <summary>
    ''' The Event Log window, made once and hidden. It is never destroyed while the
    ''' program runs, so what has been logged survives being closed.
    ''' </summary>
    Private LogWindow As EventLogWindow

    Private Sub CreateEventLogWindow()

        If LogWindow IsNot Nothing Then Exit Sub

        LogWindow = New EventLogWindow With {.Owner = Shell}

        ' Show and hide once: a window's Loaded event, and therefore its template and
        ' its grid, do not exist until it has been shown at least once.
        LogWindow.Show()
        LogWindow.Hide()

    End Sub

    ''' <summary>The Command window - made once and hidden, exactly like the log.</summary>
    Private CommandsWindow As CommandWindow

    Private Sub CreateCommandWindow()

        If CommandsWindow IsNot Nothing Then Exit Sub

        CommandsWindow = New CommandWindow With {.Owner = Shell}

        CommandsWindow.Show()
        CommandsWindow.Hide()

    End Sub

    ''' <summary>The Terminal window, on the same terms as the other two.</summary>
    Private TerminalsWindow As TerminalWindow

    Private Sub CreateTerminalWindow()

        If TerminalsWindow IsNot Nothing Then Exit Sub

        TerminalsWindow = New TerminalWindow With {.Owner = Shell}

        TerminalsWindow.Show()
        TerminalsWindow.Hide()

    End Sub

    ''' <summary>
    ''' The project name as it should appear in a file name - no path, no extension,
    ''' and never blank.
    ''' </summary>
    Public Function ProjectNameForFiles() As String

        Return ProjectTitleName()

    End Function

    ''' <summary>Asks the user to close the window. The Closing handler does the rest.</summary>
    Public Sub ExitApplication()

        If Shell Is Nothing Then Exit Sub
        Shell.Close()

    End Sub

    ''' <summary>
    ''' Everything that has to happen before the window closes. Called from
    ''' Window_Closing, so Alt+F4 and the caption's X behave like File / Exit.
    ''' </summary>
    Public Sub HandleWindowClosing(e As CancelEventArgs)

        If AppSettings.Current.ConfirmOnExit Then
            Dim answer As MessageBoxResult = MessageBox.Show(
                Shell,
                "Close " & AppTitle & "?",
                AppTitle,
                MessageBoxButton.OKCancel,
                MessageBoxImage.Question)

            If answer <> MessageBoxResult.OK Then
                e.Cancel = True
                SetStatus("Exit cancelled")
                Exit Sub
            End If
        End If

        If Not ConfirmDiscardChanges() Then
            e.Cancel = True
            SetStatus("Exit cancelled")
            Exit Sub
        End If

        StoreWindowGeometry()

        ' A serial port left open outlives the window that used it, and the next
        ' run would find it in use.
        ProbeControl.Disconnect()

    End Sub


    ' ========================================================================
    '  Project commands
    ' ========================================================================

    ''' <summary>Starts a new, unsaved project.</summary>
    Public Sub NewProject()

        If Not ConfirmDiscardChanges() Then
            SetStatus("New project cancelled")
            Exit Sub
        End If

        CurrentProject = New ProjectData With {
            .ProjectName = UntitledName,
            .CreatedUtc = UtcStamp(),
            .ModifiedUtc = UtcStamp()
        }

        CurrentProjectPath = String.Empty

        ' A new project starts on an empty bus, not on whatever the last one left
        ' lying around.
        If Shell IsNot Nothing Then Shell.stk_Devices.Children.Clear()

        ' Not dirty: an empty project has nothing to lose yet, and marking it so only
        ' produces a pointless "save first?" the moment anything else is opened.
        IsModified = False

        UpdateActiveProject()
        WorkspaceCore.Reload()

        SetStatus("New project created - not saved yet")

    End Sub

    ''' <summary>
    ''' Where .Busmaster files live, beside the executable like the settings and the
    ''' device library.
    ''' </summary>
    Public ReadOnly Property ProjectFolder As String
        Get
            Return Path.Combine(AppContext.BaseDirectory, ProjectFolderName)
        End Get
    End Property

    Private Sub EnsureProjectFolder()

        If Not Directory.Exists(ProjectFolder) Then Directory.CreateDirectory(ProjectFolder)

    End Sub

    ''' <summary>Asks for a project file and loads it.</summary>
    Public Sub OpenProject()

        If Not ConfirmDiscardChanges() Then
            SetStatus("Open cancelled")
            Exit Sub
        End If

        Dim dialog As New OpenFileDialog With {
            .Title = "Open Project",
            .Filter = ProjectFilter,
            .DefaultExt = ProjectExtension,
            .CheckFileExists = True,
            .InitialDirectory = InitialFolder()
        }

        Dim picked As Boolean? = dialog.ShowDialog(Shell)

        If Not picked.GetValueOrDefault() Then
            SetStatus("Open cancelled")
            Exit Sub
        End If

        LoadProjectFile(dialog.FileName)

    End Sub

    ''' <summary>
    ''' Loads the project behind one of the MRU menu items. The path travels on the
    ''' menu item's Tag, put there by RefreshRecentMenu.
    ''' </summary>
    Public Sub OpenRecentProject(menuSource As Object)

        Dim item As MenuItem = TryCast(menuSource, MenuItem)
        If item Is Nothing Then Exit Sub

        Dim projectPath As String = TryCast(item.Tag, String)

        If String.IsNullOrWhiteSpace(projectPath) Then
            SetStatus("That recent entry is empty")
            Exit Sub
        End If

        If Not ConfirmDiscardChanges() Then
            SetStatus("Open cancelled")
            Exit Sub
        End If

        LoadProjectFile(projectPath)

    End Sub

    ''' <summary>
    ''' Saves the workspace. A project that has never been written asks for a name
    ''' first, and only that first save goes on the recent list - after that the file
    ''' is simply overwritten, silently.
    ''' </summary>
    Public Function SaveProject() As Boolean

        Dim firstSave As Boolean = String.IsNullOrEmpty(CurrentProjectPath)
        Dim target As String

        If firstSave Then
            target = AskWhereToSave()
            If target.Length = 0 Then
                SetStatus("Save cancelled")
                Return False
            End If
        Else
            target = CurrentProjectPath
        End If

        Return WriteProjectFile(target, firstSave)

    End Function

    ''' <summary>Always asks for a name, and always goes on the recent list.</summary>
    Public Function SaveProjectAs() As Boolean

        Dim target As String = AskWhereToSave()

        If target.Length = 0 Then
            SetStatus("Save cancelled")
            Return False
        End If

        Return WriteProjectFile(target, True)

    End Function

    ''' <summary>
    ''' The name prompt. Overwriting is silent, as asked, so the dialog does not ask
    ''' either.
    ''' </summary>
    Private Function AskWhereToSave() As String

        EnsureProjectFolder()

        Dim dialog As New SaveFileDialog With {
            .Title = "Save Project",
            .Filter = ProjectFilter,
            .DefaultExt = ProjectExtension,
            .AddExtension = True,
            .OverwritePrompt = False,
            .InitialDirectory = InitialFolder(),
            .FileName = SuggestedFileName()
        }

        Dim picked As Boolean? = dialog.ShowDialog(Shell)

        If Not picked.GetValueOrDefault() Then Return String.Empty

        Return dialog.FileName

    End Function

    ''' <summary>File / Options - program-wide settings.</summary>
    Public Sub ShowOptions()

        ' TODO: dark-themed options window editing AppSettings.Current, then
        '       AppSettings.Save() and RefreshRecentMenu().
        ReportNotImplemented("Options")

    End Sub

    ''' <summary>
    ''' Status bar report for commands that are wired up but not written yet.
    ''' Remove each call as its command gets a real body.
    ''' </summary>
    Public Sub ReportNotImplemented(commandName As String)

        SetStatus(commandName & " - not implemented yet")

    End Sub

    ''' <summary>
    ''' Status bar report for a value a control refused to accept. Raised by
    ''' BitFieldEditor when an out-of-range entry is committed.
    ''' </summary>
    Public Sub ReportIllegalData()

        SetStatus("Illegal Data")

    End Sub

    ''' <summary>
    ''' Status bar report for a device file that could not be read. Raised by
    ''' DevicePanel - deliberately quiet, no dialog.
    ''' </summary>
    Public Sub ReportDeviceFileMissing()

        SetStatus("File not found")

    End Sub


    ' ========================================================================
    '  Decimal / hexadecimal
    '
    '  One switch for every register on screen, wherever it sits. Only the way the
    '  number is written changes - no value is touched, so nothing downstream sees
    '  an edit.
    ' ========================================================================

    ''' <summary>How every BitFieldEditor is currently writing its value.</summary>
    Public Property ValueDisplay As BitFieldValueType = BitFieldValueType.mode_Decimal

    Public Sub ToggleValueDisplay()

        If ValueDisplay = BitFieldValueType.mode_Hexadecimal Then
            ShowValuesAsDecimal()
        Else
            ShowValuesAsHexadecimal()
        End If

    End Sub

    Public Sub ShowValuesAsDecimal()

        SetValueDisplay(BitFieldValueType.mode_Decimal)

    End Sub

    Public Sub ShowValuesAsHexadecimal()

        SetValueDisplay(BitFieldValueType.mode_Hexadecimal)

    End Sub

    Private Sub SetValueDisplay(mode As BitFieldValueType)

        ValueDisplay = mode

        ApplyValueDisplay()
        RefreshValueDisplayControls()

    End Sub

    ''' <summary>Pushes the current radix onto every register on screen.</summary>
    Public Sub ApplyValueDisplay()

        If Shell Is Nothing Then Exit Sub

        ForEachDescendant(Of BitFieldEditor)(Shell, Sub(editor) editor.ValueType = ValueDisplay)

    End Sub

    ''' <summary>Keeps the toolbar button and the View menu showing the real state.</summary>
    Private Sub RefreshValueDisplayControls()

        If Shell Is Nothing Then Exit Sub

        Dim hexadecimal As Boolean = (ValueDisplay = BitFieldValueType.mode_Hexadecimal)

        Shell.tlbr_ValueType.Content = If(hexadecimal, "HEX", "DEC")
        Shell.tlbr_ValueType.Foreground = TryCast(
            Shell.TryFindResource(If(hexadecimal, "Brush_Value_Hexadecimal", "Brush_Value_Decimal")), Brush)

        Shell.mnu_View_Hexadecimal.IsChecked = hexadecimal
        Shell.mnu_View_Decimal.IsChecked = Not hexadecimal

    End Sub


    ' ========================================================================
    '  Devices on screen
    ' ========================================================================

    ''' <summary>Opens or closes every device at once, from the toolbar triangle.</summary>
    Public Sub SetAllDevicesExpanded(expanded As Boolean)

        If Shell Is Nothing Then Exit Sub

        ForEachDescendant(Of DevicePanel)(Shell, Sub(panel) panel.IsExpanded = expanded)

        SetStatus(If(expanded, "All devices opened", "All devices closed"))

    End Sub

    ''' <summary>
    ''' The device panels on screen, top to bottom. One place that knows how the
    ''' workspace is put together, for anything that has to visit them all.
    ''' </summary>
    Public Function DevicePanels() As List(Of DevicePanel)

        Dim panels As New List(Of DevicePanel)
        If Shell Is Nothing Then Return panels

        For Each child As UIElement In Shell.stk_Devices.Children

            Dim panel As DevicePanel = TryCast(child, DevicePanel)
            If panel IsNot Nothing Then panels.Add(panel)

        Next

        Return panels

    End Function

    ''' <summary>
    ''' Puts another device on screen, below the ones already there, and brings it
    ''' into line with the radix currently in force.
    ''' </summary>
    Public Sub AddDevicePanel(devicefile As String)

        If Shell Is Nothing Then Exit Sub

        Dim panel As DevicePanel = NewDevicePanel()

        Shell.stk_Devices.Children.Add(panel)
        panel.LoadDevice(devicefile)

        ' The editors do not exist in the tree until the panel has been laid out.
        Shell.UpdateLayout()
        ApplyValueDisplay()

        MarkModified()
        SetStatus("Added " & Path.GetFileNameWithoutExtension(devicefile))

    End Sub

    Private Function NewDevicePanel() As DevicePanel

        Return New DevicePanel With {
            .Margin = New Thickness(0, 0, 0, 6),
            .HorizontalAlignment = HorizontalAlignment.Stretch
        }

    End Function

    ''' <summary>
    ''' Right-click / Clone. Same part, same file, named after the one it came from,
    ''' and dropped in below everything else.
    ''' </summary>
    Public Sub CloneDevicePanel(source As Object)

        Dim original As DevicePanel = TryCast(source, DevicePanel)
        If original Is Nothing OrElse Shell Is Nothing Then Exit Sub

        Dim devicefile As String = If(original.DeviceFile, String.Empty)

        If devicefile.Length = 0 Then
            SetStatus("Nothing to clone - this device has no file")
            Exit Sub
        End If

        Dim copy As DevicePanel = NewDevicePanel()
        Shell.stk_Devices.Children.Add(copy)

        copy.LoadDevice(devicefile)
        copy.FunctionName = "Copy of " & If(original.FunctionName, String.Empty)

        Shell.UpdateLayout()
        ApplyValueDisplay()

        MarkModified()
        SetStatus("Cloned " & original.DeviceName)

    End Sub

    ''' <summary>Right-click / Delete. The panel goes for good.</summary>
    Public Sub RemoveDevicePanel(source As Object)

        Dim panel As DevicePanel = TryCast(source, DevicePanel)
        If panel Is Nothing OrElse Shell Is Nothing Then Exit Sub

        Dim name As String = panel.DeviceName

        Shell.stk_Devices.Children.Remove(panel)

        MarkModified()
        SetStatus("Removed " & name)

    End Sub


    ' ========================================================================
    '  Folders
    ' ========================================================================

    ''' <summary>Tools / Show Config Files - where settings.json lives.</summary>
    Public Sub ShowConfigFolder()

        OpenInExplorer(AppContext.BaseDirectory, "settings")

    End Sub

    ''' <summary>Tools / Show Dev Files.</summary>
    Public Sub ShowDeviceFolder()

        DeviceLibrary.EnsureFolder()
        OpenInExplorer(DeviceLibrary.DeviceFolder, "device")

    End Sub

    ''' <summary>Tools / Show Project Files.</summary>
    Public Sub ShowProjectFolder()

        EnsureProjectFolder()
        OpenInExplorer(ProjectFolder, "project")

    End Sub

    Private Sub OpenInExplorer(folder As String, description As String)

        Try
            If Not Directory.Exists(folder) Then
                SetStatus("No " & description & " folder yet")
                Exit Sub
            End If

            Process.Start(New ProcessStartInfo(folder) With {.UseShellExecute = True})
            SetStatus("Opened the " & description & " folder")

        Catch ex As Exception
            SetStatus("Could not open the " & description & " folder - " & ex.Message)
        End Try

    End Sub

    ''' <summary>Walks the visual tree and hands every match to the caller.</summary>
    Private Sub ForEachDescendant(Of T As DependencyObject)(root As DependencyObject, action As Action(Of T))

        If root Is Nothing Then Exit Sub

        Dim children As Integer = VisualTreeHelper.GetChildrenCount(root)

        For index As Integer = 0 To children - 1

            Dim child As DependencyObject = VisualTreeHelper.GetChild(root, index)

            Dim match As T = TryCast(child, T)
            If match IsNot Nothing Then action(match)

            ForEachDescendant(Of T)(child, action)

        Next

    End Sub


    ' ========================================================================
    '  Operation mode
    '
    '  When a change made on screen goes out to the bus. Recorded here; the bus
    '  code will read it once there is any.
    ' ========================================================================

    ''' <summary>How the toolbar's Operation list is currently set.</summary>
    Public Property OperationMode As BusOperationMode = BusOperationMode.mode_RegisterOnChange

    ''' <summary>Called with the index of the item the user picked.</summary>
    Public Sub OperationModeChosen(selectedIndex As Integer)

        Select Case selectedIndex
            Case 1
                OperationMode = BusOperationMode.mode_DeviceOnChange
                SetStatus("Operation: write the whole device when anything changes")
            Case Else
                OperationMode = BusOperationMode.mode_RegisterOnChange
                SetStatus("Operation: write a register when it changes")
        End Select

    End Sub


    ' ========================================================================
    '  The probe
    '
    '  The port list and the connect toggle. The toggle shows whether the link is
    '  actually up, not what was last asked for, so a port that refuses to open
    '  puts it back where it was.
    ' ========================================================================

    ''' <summary>Guards the port list while it is being refilled.</summary>
    Private RefillingProbeList As Boolean = False

    ''' <summary>
    ''' Fills the port list from what the machine has now, and marks the one the
    ''' settings name. The blank first entry is "no probe", which is what leaves the
    ''' Terminal turning lines round on its own.
    ''' </summary>
    Public Sub RefreshProbePorts()

        FillProbeList(ProbeControl.AvailablePorts())

    End Sub

    ''' <summary>
    ''' Toolbar / Probe / rescan. Asks the USB bus which serial ports belong to a
    ''' BusMaster probe and offers those. With none found the list falls back to
    ''' every serial port, so a probe that does not announce itself properly can
    ''' still be reached by hand.
    ''' </summary>
    Public Sub RescanProbes()

        Dim probes As List(Of ProbePort) = ProbeControl.FindProbePorts()

        If probes.Count > 0 Then

            FillProbeList(probes)
            SetStatus(Describe(probes.Count) & " found on " &
                      String.Join(", ", probes.Select(Function(p) p.PortName)))

        Else

            FillProbeList(ProbeControl.AvailablePorts())
            SetStatus("No BusMaster probe found - listing every serial port")

        End If

    End Sub

    ''' <summary>"1 probe" or "3 probes" - the status bar should read properly.</summary>
    Private Function Describe(count As Integer) As String

        If count = 1 Then Return "1 probe"

        Return count.ToString(CultureInfo.InvariantCulture) & " probes"

    End Function

    Private Sub FillProbeList(ports As IEnumerable(Of ProbePort))

        If Shell Is Nothing Then Exit Sub

        Dim list As ComboBox = Shell.tlbr_Probe_List

        RefillingProbeList = True

        list.Items.Clear()

        ' The blank first entry is "no probe", which is what leaves the Terminal
        ' turning lines round on its own.
        list.Items.Add(New ProbePort(String.Empty, String.Empty))

        For Each port As ProbePort In ports
            list.Items.Add(port)
        Next

        Dim wanted As String = If(AppSettings.Current.ProbePort, String.Empty).Trim()

        ' A port named in the settings but not plugged in still belongs on the list,
        ' or choosing it again would be impossible once it came back.
        If wanted.Length > 0 AndAlso Not ports.Any(Function(p) Same(p.PortName, wanted)) Then
            list.Items.Add(New ProbePort(wanted, String.Empty))
        End If

        list.SelectedItem = list.Items.OfType(Of ProbePort)().
                                 FirstOrDefault(Function(p) Same(p.PortName, wanted))

        RefillingProbeList = False

    End Sub

    Private Function Same(left As String, right As String) As Boolean

        Return String.Equals(left, right, StringComparison.OrdinalIgnoreCase)

    End Function

    ''' <summary>The user picked a port from the list.</summary>
    Public Sub ProbePortChosen(chosen As Object)

        If RefillingProbeList Then Exit Sub

        Dim port As ProbePort = TryCast(chosen, ProbePort)

        ProbeControl.ChoosePort(If(port Is Nothing, String.Empty, port.PortName))
        RefreshProbeConnection()

    End Sub

    ''' <summary>The connect toggle was pressed one way or the other.</summary>
    Public Sub ProbeConnectionRequested(connect As Boolean)

        If RefillingProbeList Then Exit Sub

        If connect Then
            ProbeControl.Connect()
        Else
            ProbeControl.Disconnect()
            SetStatus("Probe disconnected")
        End If

        RefreshProbeConnection()

    End Sub

    ''' <summary>
    ''' Puts the toggle where the link actually is. Called after anything that could
    ''' have changed it, so a failed connection does not leave the button claiming
    ''' otherwise.
    ''' </summary>
    Public Sub RefreshProbeConnection()

        If Shell Is Nothing Then Exit Sub

        Dim live As Boolean = ProbeControl.IsConnected

        ' Which port the probe is on cannot be changed out from under an open one,
        ' and there is no point offering to look for another while it is.
        Shell.tlbr_Probe_List.IsEnabled = Not live
        Shell.tlbr_Probe_Rescan.IsEnabled = Not live

        ' Everything to the right of the connect button drives a wire, so none of
        ' it means anything with no probe on the other end. Two panels, so this is
        ' two lines rather than a list that has to be kept up to date as controls
        ' are added.
        Shell.tlbr_Bus_Block.IsEnabled = live
        Shell.tlbr_Target_Block.IsEnabled = live

        If Shell.tlbr_Probe_Connect.IsChecked.GetValueOrDefault() = live Then Exit Sub

        RefillingProbeList = True
        Shell.tlbr_Probe_Connect.IsChecked = live
        RefillingProbeList = False

    End Sub


    ' ========================================================================
    '  Recording
    '
    '  Whether what happens on screen is written to the event log. The toolbar
    '  button holds the state itself, so there is only ever one copy of it and it
    '  cannot drift away from what the user can see.
    ' ========================================================================

    ''' <summary>True while the toolbar's record button shows a red disc.</summary>
    Public ReadOnly Property IsRecording As Boolean
        Get
            If Shell Is Nothing Then Return False

            Return Shell.tlbr_Record.IsChecked.GetValueOrDefault()
        End Get
    End Property

    ''' <summary>Called when the record / pause button is clicked either way.</summary>
    Public Sub RecordingChanged(recording As Boolean)

        If recording Then
            SetStatus("Recording to the event log")
        Else
            SetStatus("Recording paused")
        End If

    End Sub


    ''' <summary>
    ''' Shows the help box.
    '''
    ''' The mouse section is the part that earns its keep: a keyboard shortcut is
    ''' written on the menu beside its command, but nothing on screen says that a
    ''' right-click reads a register. Anything that cannot be found by looking at
    ''' the window belongs here.
    ''' </summary>
    Public Sub ShowHelp()

        ' Written as sentences rather than in columns: a message box uses whatever
        ' font the system gives it, and padded columns come out ragged in every
        ' proportional one.
        Dim text As String =
            AppTitle & " - I2C bus mastering controller" & vbCrLf & vbCrLf &
            "Keyboard" & vbCrLf &
            "   Ctrl+N - New Project" & vbCrLf &
            "   Ctrl+O - Open Project" & vbCrLf &
            "   Ctrl+S - Save Project" & vbCrLf &
            "   Ctrl+Shift+S - Save Project As" & vbCrLf &
            "   F1 - this help" & vbCrLf & vbCrLf &
            "Registers" & vbCrLf &
            "   Left-click a bit - toggle it, and write it out" & vbCrLf &
            "   Right-click the bits - read that register" & vbCrLf &
            "   Ctrl+right-click the bits - read every register on the device" & vbCrLf &
            "   Padlock - keep the register on show when the device is closed" & vbCrLf & vbCrLf &
            "Toolbar" & vbCrLf &
            "   Operation - whether a change writes just that register, or the" & vbCrLf &
            "      whole device it belongs to" & vbCrLf &
            "   Write All - write every register on every device" & vbCrLf &
            "   Read All - read every register on every device" & vbCrLf & vbCrLf &
            "Device header" & vbCrLf &
            "   Click the name - rename it" & vbCrLf &
            "   Drag the bars - move the device up or down the stack" & vbCrLf &
            "   Right-click - clone or delete the device" & vbCrLf &
            "   Triangle - open or close the device" & vbCrLf & vbCrLf &
            "Workspace views" & vbCrLf &
            "   Save - save the padlocks as a new view" & vbCrLf &
            "   Right-click Save - update the view showing in the list" & vbCrLf &
            "   Delete - throw that view away" & vbCrLf & vbCrLf &
            "Event log" & vbCrLf &
            "   Record - log every register the user changes" & vbCrLf &
            "   Double-click a line number - set or clear a breakpoint" & vbCrLf &
            "   L S P D / - marker, typed in the second column" & vbCrLf & vbCrLf &
            "Settings, options and the recent project list are kept in:" & vbCrLf &
            AppSettings.SettingsFilePath

        MessageBox.Show(Shell, text, AppTitle & " Help", MessageBoxButton.OK, MessageBoxImage.Information)
        SetStatus("Help shown")

    End Sub

    ''' <summary>
    ''' Flags the open project as changed. Call this from wherever project data is
    ''' edited, so the unsaved-changes prompt and the "*" marker work.
    ''' </summary>
    Public Sub MarkModified()

        IsModified = True
        UpdateActiveProject()

    End Sub

    ''' <summary>
    ''' The workspace backdrop: a two-tone checkerboard tiled from the sizes and
    ''' colours in AppSettings.
    ''' </summary>
    Private Sub ApplyWorkspaceBackground()

        If Shell Is Nothing Then Exit Sub

        Dim side As Double = AppSettings.Checker_Size
        Dim tile As Double = side * 2

        Dim light As New SolidColorBrush(ParseColour(AppSettings.Checker_Light))
        Dim dark As New SolidColorBrush(ParseColour(AppSettings.Checker_Dark))

        Dim squares As New DrawingGroup()

        ' Dark fills the tile; light takes the two diagonal squares.
        squares.Children.Add(New GeometryDrawing(dark, Nothing,
            New RectangleGeometry(New Rect(0, 0, tile, tile))))
        squares.Children.Add(New GeometryDrawing(light, Nothing,
            New RectangleGeometry(New Rect(0, 0, side, side))))
        squares.Children.Add(New GeometryDrawing(light, Nothing,
            New RectangleGeometry(New Rect(side, side, side, side))))

        Dim checker As New DrawingBrush(squares) With {
            .TileMode = TileMode.Tile,
            .Viewport = New Rect(0, 0, tile, tile),
            .ViewportUnits = BrushMappingMode.Absolute,
            .Stretch = Stretch.None
        }

        checker.Freeze()

        Shell.scr_Workspace.Background = checker

    End Sub

    Private Function ParseColour(value As String) As Color

        Try
            Return CType(ColorConverter.ConvertFromString(value), Color)
        Catch
            Return Colors.Black
        End Try

    End Function


    ' ========================================================================
    '  Project file I/O
    ' ========================================================================

    ''' <summary>Reads a project file and makes it the open project.</summary>
    Public Sub LoadProjectFile(projectPath As String)

        If String.IsNullOrWhiteSpace(projectPath) Then
            SetStatus("No project file specified")
            Exit Sub
        End If

        If Not File.Exists(projectPath) Then
            AppSettings.RemoveRecentProject(projectPath)
            RefreshRecentMenu()
            SetStatus("Project not found - " & projectPath)
            MessageBox.Show(Shell,
                            "This project file no longer exists:" & vbCrLf & vbCrLf & projectPath,
                            AppTitle, MessageBoxButton.OK, MessageBoxImage.Warning)
            Exit Sub
        End If

        Try
            Dim loaded As ProjectData = JsonFile.ReadFrom(Of ProjectData)(projectPath)

            CurrentProject = loaded
            CurrentProjectPath = Path.GetFullPath(projectPath)
            IsModified = False

            If String.IsNullOrWhiteSpace(CurrentProject.ProjectName) Then
                CurrentProject.ProjectName = Path.GetFileNameWithoutExtension(CurrentProjectPath)
            End If

            RebuildWorkspace(CurrentProject)

            AppSettings.Current.LastProjectFolder = Path.GetDirectoryName(CurrentProjectPath)
            AppSettings.AddRecentProject(CurrentProjectPath)

            RefreshRecentMenu()
            UpdateActiveProject()
            WorkspaceCore.Reload()

            SetStatus("Opened " & Path.GetFileName(CurrentProjectPath) &
                      " (" & CurrentProject.Devices.Count.ToString(CultureInfo.InvariantCulture) & " devices)")

        Catch ex As Exception
            SetStatus("Open failed - " & ex.Message)
            MessageBox.Show(Shell,
                            "The project could not be opened:" & vbCrLf & vbCrLf & ex.Message,
                            AppTitle, MessageBoxButton.OK, MessageBoxImage.Error)
        End Try

    End Sub

    ''' <summary>
    ''' Writes the workspace out. addToRecent is False for a plain re-save, which is
    ''' why an established project does not keep bubbling to the top of the MRU list.
    ''' </summary>
    Private Function WriteProjectFile(projectPath As String, addToRecent As Boolean) As Boolean

        Try
            Dim fullPath As String = Path.GetFullPath(projectPath)

            If CurrentProject Is Nothing Then CurrentProject = New ProjectData

            ' A project that was never named takes its name from the file.
            If String.IsNullOrWhiteSpace(CurrentProject.ProjectName) OrElse
               CurrentProject.ProjectName = UntitledName Then
                CurrentProject.ProjectName = Path.GetFileNameWithoutExtension(fullPath)
            End If

            If String.IsNullOrWhiteSpace(CurrentProject.CreatedUtc) Then
                CurrentProject.CreatedUtc = UtcStamp()
            End If

            CurrentProject.ModifiedUtc = UtcStamp()
            CurrentProject.Devices = CollectWorkspace()

            JsonFile.WriteTo(fullPath, CurrentProject)

            CurrentProjectPath = fullPath
            IsModified = False

            AppSettings.Current.LastProjectFolder = Path.GetDirectoryName(fullPath)

            If addToRecent Then
                AppSettings.AddRecentProject(fullPath)
                RefreshRecentMenu()
            End If

            UpdateActiveProject()

            ' Views are named after the project file, so a first save or a Save As
            ' changes which set of them belongs to what is on screen.
            WorkspaceCore.Reload()

            SetStatus("Saved " & Path.GetFileName(fullPath) &
                      " (" & CurrentProject.Devices.Count.ToString(CultureInfo.InvariantCulture) & " devices)")
            Return True

        Catch ex As Exception
            SetStatus("Save failed - " & ex.Message)
            MessageBox.Show(Shell,
                            "The project could not be saved:" & vbCrLf & vbCrLf & ex.Message,
                            AppTitle, MessageBoxButton.OK, MessageBoxImage.Error)
            Return False
        End Try

    End Function

    ''' <summary>
    ''' The devices on screen, top to bottom. Only the function name and which part
    ''' it is - enough to build the layout again, nothing more.
    ''' </summary>
    Private Function CollectWorkspace() As List(Of ProjectDeviceData)

        Dim devices As New List(Of ProjectDeviceData)
        If Shell Is Nothing Then Return devices

        For Each child As UIElement In Shell.stk_Devices.Children

            Dim panel As DevicePanel = TryCast(child, DevicePanel)
            If panel Is Nothing Then Continue For

            devices.Add(New ProjectDeviceData With {
                .PanelId = If(panel.PanelId, String.Empty),
                .FunctionName = If(panel.FunctionName, String.Empty),
                .DeviceFile = Path.GetFileNameWithoutExtension(If(panel.DeviceFile, String.Empty))
            })

        Next

        Return devices

    End Function

    ''' <summary>
    ''' Throws away what is on screen and builds the saved layout in its place, in
    ''' the order the file lists it.
    ''' </summary>
    Private Sub RebuildWorkspace(data As ProjectData)

        If Shell Is Nothing Then Exit Sub

        Shell.stk_Devices.Children.Clear()

        ' Set if any panel had to be given an identity the file did not carry.
        Dim invented As Boolean = False

        If data.Devices IsNot Nothing Then

            For Each device As ProjectDeviceData In data.Devices

                If device Is Nothing Then Continue For

                Dim panel As DevicePanel = NewDevicePanel()
                Shell.stk_Devices.Children.Add(panel)

                If String.IsNullOrWhiteSpace(device.PanelId) Then

                    ' A project written before panels had identities. The one the
                    ' panel made for itself goes back into the project, because a
                    ' view saved this session would not find its way home next time
                    ' unless the file learns it.
                    device.PanelId = panel.PanelId
                    invented = True

                Else
                    panel.PanelId = device.PanelId
                End If

                panel.LoadDevice(If(device.DeviceFile, String.Empty))
                panel.FunctionName = If(device.FunctionName, String.Empty)

            Next

        End If

        Shell.UpdateLayout()
        ApplyValueDisplay()

        ' The project in hand now says something the file does not. Marking it so is
        ' honest, and it is what gets those identities written down - a one-off, the
        ' first time an older project is opened.
        If invented Then MarkModified()

    End Sub

    ''' <summary>
    ''' Yes / No / Cancel prompt guarding any action that would throw away unsaved
    ''' work. Returns True when the caller may proceed.
    ''' </summary>
    Private Function ConfirmDiscardChanges() As Boolean

        ' Deliberately does not test CurrentProject: devices can be added to the
        ' workspace before any project exists, and those changes are worth keeping.
        If Not IsModified Then Return True

        Dim answer As MessageBoxResult = MessageBox.Show(
            Shell,
            "Save changes to " & ProjectDisplayName() & " first?",
            AppTitle,
            MessageBoxButton.YesNoCancel,
            MessageBoxImage.Question)

        Select Case answer
            Case MessageBoxResult.Yes
                Return SaveProject()
            Case MessageBoxResult.No
                Return True
            Case Else
                Return False
        End Select

    End Function


    ' ========================================================================
    '  User interface updates
    ' ========================================================================

    ''' <summary>Writes to sbar_status - the result of the last operation.</summary>
    Public Sub SetStatus(message As String)

        If Shell Is Nothing Then Exit Sub

        Shell.sbar_status.Content = If(String.IsNullOrEmpty(message), "Ready", message)

    End Sub

    ''' <summary>
    ''' Writes to sbar_ActiveProject - the open project's name and full path - and
    ''' keeps the window title in step.
    ''' </summary>
    Private Sub UpdateActiveProject()

        If Shell Is Nothing Then Exit Sub

        Dim caption As String

        If CurrentProject Is Nothing Then
            caption = "No project"
        ElseIf String.IsNullOrEmpty(CurrentProjectPath) Then
            caption = CurrentProject.ProjectName & "   (not saved yet)"
        Else
            caption = CurrentProject.ProjectName & "   -   " & CurrentProjectPath
        End If

        If IsModified Then caption &= "   *"

        Shell.sbar_ActiveProject.Content = caption

        ' The status bar carries the full path; the title bar only needs the name.
        Dim title As String = AppName & " (" & ProjectTitleName() & ")  Version " &
                              AppSettings.ApplicationVersion

        If IsModified Then title &= " *"

        Shell.Title = title

    End Sub

    ''' <summary>
    ''' The project's name for the title bar: the file name with the path and the
    ''' .Busmaster extension taken off. Falls back to the in-memory name while the
    ''' project has never been written.
    ''' </summary>
    Private Function ProjectTitleName() As String

        If Not String.IsNullOrEmpty(CurrentProjectPath) Then
            Return Path.GetFileNameWithoutExtension(CurrentProjectPath)
        End If

        Return ProjectDisplayName()

    End Function

    ''' <summary>
    ''' Fills mnu_File_Recent_1 .. mnu_File_Recent_5 from the settings file, hiding
    ''' the slots that have no entry.
    ''' </summary>
    Public Sub RefreshRecentMenu()

        If Shell Is Nothing Then Exit Sub

        Dim slots() As MenuItem = {Shell.mnu_File_Recent_1,
                                   Shell.mnu_File_Recent_2,
                                   Shell.mnu_File_Recent_3,
                                   Shell.mnu_File_Recent_4,
                                   Shell.mnu_File_Recent_5}

        Dim recent As List(Of String) = AppSettings.Current.RecentProjects

        For index As Integer = 0 To slots.Length - 1

            Dim slot As MenuItem = slots(index)

            If index < recent.Count Then
                slot.Header = "_" & (index + 1).ToString() & "   " & EscapeAccessKeys(recent(index))
                slot.Tag = recent(index)
                slot.ToolTip = recent(index)
                slot.Visibility = Visibility.Visible
            Else
                slot.Header = String.Empty
                slot.Tag = Nothing
                slot.ToolTip = Nothing
                slot.Visibility = Visibility.Collapsed
            End If

        Next

        Shell.mnu_File_Recent_Empty.Visibility =
            If(recent.Count = 0, Visibility.Visible, Visibility.Collapsed)

    End Sub

    Private Sub RestoreWindowGeometry()

        Shell.Width = AppSettings.Current.WindowWidth
        Shell.Height = AppSettings.Current.WindowHeight

        If AppSettings.Current.WindowMaximized Then Shell.WindowState = WindowState.Maximized

    End Sub

    Private Sub StoreWindowGeometry()

        AppSettings.Current.WindowMaximized = (Shell.WindowState = WindowState.Maximized)

        ' Only a normal window reports the size worth restoring.
        If Shell.WindowState = WindowState.Normal Then
            AppSettings.Current.WindowWidth = Shell.ActualWidth
            AppSettings.Current.WindowHeight = Shell.ActualHeight
        End If

        AppSettings.Save()

    End Sub


    ' ========================================================================
    '  Small helpers
    ' ========================================================================

    ''' <summary>Sortable, culture-independent UTC timestamp for the project file.</summary>
    Private Function UtcStamp() As String
        Return DateTime.UtcNow.ToString("s") & "Z"
    End Function

    ''' <summary>Folder the file dialogs should open in.</summary>
    Private Function InitialFolder() As String

        Dim folder As String = AppSettings.Current.LastProjectFolder

        If Not String.IsNullOrWhiteSpace(folder) AndAlso Directory.Exists(folder) Then Return folder

        EnsureProjectFolder()
        Return ProjectFolder

    End Function

    ''' <summary>File name to offer in the Save As dialog.</summary>
    Private Function SuggestedFileName() As String

        If Not String.IsNullOrEmpty(CurrentProjectPath) Then Return Path.GetFileName(CurrentProjectPath)

        ' The workspace can be saved without a project ever having been created, so
        ' there may be no CurrentProject to take a name from.
        Return ProjectDisplayName() & ProjectExtension

    End Function

    ''' <summary>The open project's name, or a stand-in when there is no project object.</summary>
    Private Function ProjectDisplayName() As String

        If CurrentProject Is Nothing Then Return UntitledName

        If String.IsNullOrWhiteSpace(CurrentProject.ProjectName) Then Return UntitledName

        Return CurrentProject.ProjectName

    End Function

    ''' <summary>
    ''' Doubles underscores so a path like C:\my_projects\a.bmproj does not turn its
    ''' underscore into a menu access key.
    ''' </summary>
    Private Function EscapeAccessKeys(text As String) As String

        If String.IsNullOrEmpty(text) Then Return String.Empty

        Return text.Replace("_", "__")

    End Function

End Module
