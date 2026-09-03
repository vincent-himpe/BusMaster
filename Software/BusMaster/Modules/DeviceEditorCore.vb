' ============================================================================
'  Modules\DeviceEditorCore.vb
'
'  All of the Device Editor's active code. The window's handlers only call in
'  here.
'
'  The three modes differ in exactly three places:
'
'      mode_New     name is editable, Save asks before overwriting
'      mode_Edit    name is locked, Save overwrites without asking
'      mode_Clone   name is editable and flagged red until changed, and Save
'                   refuses outright if that name is already taken
'
'  The editor is always shown with ShowDialog, one at a time, so a single module
'  level reference to the open window is enough.
' ============================================================================

Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.ComponentModel
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Windows.Threading

Public Module DeviceEditorCore

    Private Const DialogTitle As String = "Device Editor"

    Private Const MinRegisterAddress As Integer = -1
    Private Const MaxRegisterAddress As Integer = 255

    ''' <summary>Display index of the Register Name column - where a new row lands.</summary>
    Private Const RegisterNameColumn As Integer = 1

    ''' <summary>Display index of D7. D0 is seven columns further along.</summary>
    Private Const FirstBitColumn As Integer = 2

    Private Const BitColumnCount As Integer = 8


    Private Editor As DeviceEditorWindow
    Private Rows As ObservableCollection(Of DeviceRegisterRow)

    ''' <summary>Name the file was loaded under. Clone mode compares against it.</summary>
    Private OriginalDeviceName As String = String.Empty

    ''' <summary>Set just before a deliberate Close so Closing does not re-ask.</summary>
    Private AllowClose As Boolean = False

    ''' <summary>Guards the device name box while its own handler rewrites it.</summary>
    Private SuppressNameSync As Boolean = False

    ''' <summary>True while Attach is filling the form, so loading is not an edit.</summary>
    Private Loading As Boolean = False

    ''' <summary>True once the user has actually changed something.</summary>
    Private IsDirty As Boolean = False

    ''' <summary>Window title without the dirty marker.</summary>
    Private BaseTitle As String = DialogTitle


    ' ========================================================================
    '  Start-up and shutdown
    ' ========================================================================

    Public Sub Attach(window As DeviceEditorWindow)

        Editor = window
        AllowClose = False
        SuppressNameSync = False
        OriginalDeviceName = String.Empty

        ' Nothing that happens while the form is being filled counts as an edit.
        Loading = True
        IsDirty = False

        WindowTheme.ApplyDarkTitleBar(Editor)

        Rows = New ObservableCollection(Of DeviceRegisterRow)
        AddHandler Rows.CollectionChanged, AddressOf Rows_CollectionChanged
        Editor.grd_Registers.ItemsSource = Rows

        AttachFieldGuards()

        Select Case Editor.EditorMode

            Case DeviceEditorMode.mode_Edit
                If LoadIntoForm(Editor.DeviceFilePath) Then
                    LockDeviceName()
                    BaseTitle = DialogTitle & " - " & OriginalDeviceName
                End If

            Case DeviceEditorMode.mode_Clone
                If LoadIntoForm(Editor.DeviceFilePath) Then
                    BaseTitle = DialogTitle & " - clone of " & OriginalDeviceName
                    ' Flag the name as needing a change, and check it on every keystroke.
                    FlagDeviceNameForRename()
                End If

            Case Else
                BaseTitle = DialogTitle & " - new device"
                Editor.txt_DeviceName.Focus()

        End Select

        RefreshAddressReadout()

        ' The grid has no new-row placeholder, so there has to be a row to type in.
        If Rows.Count = 0 Then Rows.Add(New DeviceRegisterRow)

        Loading = False
        IsDirty = False
        RefreshTitle()

    End Sub


    ' ========================================================================
    '  Clean / dirty
    ' ========================================================================

    ''' <summary>
    ''' Notes that the user has changed something. Ignored while the form is being
    ''' filled in, and only does the work once.
    ''' </summary>
    Private Sub MarkDirty()

        If Loading OrElse IsDirty OrElse Editor Is Nothing Then Exit Sub

        IsDirty = True
        RefreshTitle()

    End Sub

    Private Sub RefreshTitle()

        If Editor Is Nothing Then Exit Sub

        Editor.Title = BaseTitle & If(IsDirty, " (*)", String.Empty)

    End Sub

    ''' <summary>Device Description was typed in. Nothing to check, just an edit.</summary>
    Public Sub DescriptionEdited()

        MarkDirty()

    End Sub

    ''' <summary>
    ''' Rows added or removed. Each row is watched individually so that tabbing
    ''' through a cell without changing it does not count as an edit.
    ''' </summary>
    Private Sub Rows_CollectionChanged(sender As Object, e As NotifyCollectionChangedEventArgs)

        If e.OldItems IsNot Nothing Then
            For Each item As Object In e.OldItems
                Dim row As DeviceRegisterRow = TryCast(item, DeviceRegisterRow)
                If row IsNot Nothing Then RemoveHandler row.PropertyChanged, AddressOf Row_PropertyChanged
            Next
        End If

        If e.NewItems IsNot Nothing Then
            For Each item As Object In e.NewItems
                Dim row As DeviceRegisterRow = TryCast(item, DeviceRegisterRow)
                If row IsNot Nothing Then AddHandler row.PropertyChanged, AddressOf Row_PropertyChanged
            Next
        End If

        If e.Action <> NotifyCollectionChangedAction.Reset Then MarkDirty()

    End Sub

    Private Sub Row_PropertyChanged(sender As Object, e As PropertyChangedEventArgs)

        MarkDirty()

    End Sub


    ' ========================================================================
    '  Field guards
    '
    '  Wired up here rather than in XAML so the window file stays free of code.
    '  Each guard rejects the keystroke before it lands, and the matching paste
    '  handler stops the clipboard being used to get round it.
    ' ========================================================================

    Private Sub AttachFieldGuards()

        AddHandler Editor.txt_DeviceName.PreviewTextInput, AddressOf DeviceName_PreviewTextInput
        AddHandler Editor.txt_DeviceName.PreviewKeyDown, AddressOf NoSpace_PreviewKeyDown
        DataObject.AddPastingHandler(Editor.txt_DeviceName, AddressOf DeviceName_Pasting)

        AddHandler Editor.txt_BaseAddress.PreviewTextInput, AddressOf Digits_PreviewTextInput
        AddHandler Editor.txt_BaseAddress.PreviewKeyDown, AddressOf NoSpace_PreviewKeyDown
        DataObject.AddPastingHandler(Editor.txt_BaseAddress, AddressOf Digits_Pasting)

        AddHandler Editor.txt_AddressBits.PreviewTextInput, AddressOf Digits_PreviewTextInput
        AddHandler Editor.txt_AddressBits.PreviewKeyDown, AddressOf NoSpace_PreviewKeyDown
        DataObject.AddPastingHandler(Editor.txt_AddressBits, AddressOf Digits_Pasting)

        AddHandler Editor.grd_Registers.PreviewKeyDown, AddressOf Registers_PreviewKeyDown
        AddHandler Editor.grd_Registers.Sorting, AddressOf Registers_Sorting

    End Sub


    ' ========================================================================
    '  Grid navigation and row management
    '
    '  The table is filled in a row at a time, so it moves sideways rather than
    '  downwards:
    '
    '      Enter          next cell to the right, and off the end of the last row
    '                     it adds a new one. On an empty D7..D0 cell it first drops
    '                     that column's bit number in, so holding Enter across the
    '                     bit columns fills them 7 6 5 4 3 2 1 0.
    '      Ctrl+Enter     add a row
    '      Ctrl+Insert    add a row
    '      Ctrl+Delete    remove the current row, never the last one left
    '      Right          next cell, but only once the caret is past the last
    '                     character - otherwise it walks through the text
    '      Left           previous cell, but only from in front of the first character
    '      Home / End     first and last cell of the row
    '      Up / Down      left alone, so they still move between rows
    '
    '  The grid's own new-row placeholder is switched off in XAML: rows are only
    '  ever added here, always at the bottom, and there is always at least one.
    ' ========================================================================

    Private Sub Registers_PreviewKeyDown(sender As Object, e As KeyEventArgs)

        Dim editing As TextBox = TryCast(Keyboard.FocusedElement, TextBox)
        Dim control As Boolean = (Keyboard.Modifiers And ModifierKeys.Control) = ModifierKeys.Control

        Select Case e.Key

            Case Key.Enter
                CapitaliseRegisterName()
                If control Then
                    AddRegisterRow()
                ElseIf FillBlankBitCell() Then
                    ' Filling was this keystroke's job. Step on to the next bit so
                    ' holding Enter walks D7 down to D0, but do not also grow a row.
                    MoveCell(1)
                ElseIf AtLastCellOfLastRow() Then
                    AddRegisterRow()
                Else
                    MoveCell(1)
                End If
                e.Handled = True

            Case Key.Insert
                If control Then
                    AddRegisterRow()
                    e.Handled = True
                End If

            Case Key.Delete
                If control Then
                    RemoveCurrentRow()
                    e.Handled = True
                End If

            Case Key.Right
                If editing IsNot Nothing AndAlso CaretAtEnd(editing) Then
                    MoveCell(1)
                    e.Handled = True
                End If

            Case Key.Left
                If editing IsNot Nothing AndAlso CaretAtStart(editing) Then
                    MoveCell(-1)
                    e.Handled = True
                End If

            Case Key.Home
                MoveToRowEdge(False)
                e.Handled = True

            Case Key.End
                MoveToRowEdge(True)
                e.Handled = True

        End Select

    End Sub

    ''' <summary>
    ''' Enter in the Register Name column tidies the name up: "open drain" becomes
    ''' "Open Drain".
    ''' </summary>
    Private Sub CapitaliseRegisterName()

        Dim current As DataGridCellInfo = Editor.grd_Registers.CurrentCell
        If Not current.IsValid OrElse current.Column Is Nothing Then Exit Sub
        If current.Column.DisplayIndex <> RegisterNameColumn Then Exit Sub

        Dim row As DeviceRegisterRow = TryCast(current.Item, DeviceRegisterRow)
        If row Is Nothing Then Exit Sub

        Dim editing As TextBox = TryCast(Keyboard.FocusedElement, TextBox)
        Dim shown As String = If(editing IsNot Nothing, editing.Text, row.RegisterName)

        Dim tidied As String = TextTools.CapitaliseWords(shown)
        If String.Equals(tidied, shown, StringComparison.Ordinal) Then Exit Sub

        If editing IsNot Nothing Then
            editing.Text = tidied
            editing.CaretIndex = tidied.Length
        Else
            row.RegisterName = tidied
        End If

    End Sub

    ''' <summary>
    ''' Enter on an empty D7..D0 cell writes that column's bit number into it - 7 in
    ''' D7, 6 in D6, down to 0 in D0 - so a row of bit fields fills by holding Enter.
    '''
    ''' Only ever fills a blank cell; anything already typed is left alone. Returns
    ''' True when it wrote something.
    ''' </summary>
    Private Function FillBlankBitCell() As Boolean

        Dim current As DataGridCellInfo = Editor.grd_Registers.CurrentCell
        If Not current.IsValid OrElse current.Column Is Nothing Then Return False

        Dim position As Integer = current.Column.DisplayIndex - FirstBitColumn
        If position < 0 OrElse position > BitColumnCount - 1 Then Return False

        Dim row As DeviceRegisterRow = TryCast(current.Item, DeviceRegisterRow)
        If row Is Nothing Then Return False

        ' While a cell is open its TextBox holds the live text, which may not have
        ' reached the row yet.
        Dim editing As TextBox = TryCast(Keyboard.FocusedElement, TextBox)
        Dim shown As String = If(editing IsNot Nothing, editing.Text, row.BitName(position))

        If Not String.IsNullOrWhiteSpace(shown) Then Return False

        ' Leftmost bit column is D7, so the number counts down from there.
        Dim bit As Integer = (BitColumnCount - 1) - position
        Dim text As String = bit.ToString(CultureInfo.InvariantCulture)

        If editing IsNot Nothing Then
            editing.Text = text
            editing.CaretIndex = text.Length
        Else
            row.BitName(position) = text
        End If

        Return True

    End Function

    ''' <summary>Caret sitting after the last character, with nothing selected.</summary>
    Private Function CaretAtEnd(box As TextBox) As Boolean

        Return box.SelectionLength = 0 AndAlso box.CaretIndex >= box.Text.Length

    End Function

    ''' <summary>Caret sitting in front of the first character, with nothing selected.</summary>
    Private Function CaretAtStart(box As TextBox) As Boolean

        Return box.SelectionLength = 0 AndAlso box.CaretIndex = 0

    End Function

    ''' <summary>
    ''' Steps one cell along the current row. Stops at either end rather than
    ''' wrapping onto another row.
    ''' </summary>
    Private Sub MoveCell(offset As Integer)

        Dim grid As DataGrid = Editor.grd_Registers
        Dim current As DataGridCellInfo = grid.CurrentCell

        If Not current.IsValid OrElse current.Column Is Nothing Then Exit Sub

        Dim target As Integer = current.Column.DisplayIndex + offset

        If target < 0 Then target = 0
        If target > grid.Columns.Count - 1 Then target = grid.Columns.Count - 1

        FocusCell(current.Item, target)

    End Sub

    Private Sub MoveToRowEdge(toEnd As Boolean)

        Dim grid As DataGrid = Editor.grd_Registers
        Dim current As DataGridCellInfo = grid.CurrentCell

        If Not current.IsValid Then Exit Sub

        FocusCell(current.Item, If(toEnd, grid.Columns.Count - 1, 0))

    End Sub

    ''' <summary>
    ''' Commits whatever is being typed, then moves to a cell on the same row and
    ''' opens it for editing so typing can carry straight on.
    ''' </summary>
    Private Sub FocusCell(rowItem As Object, displayIndex As Integer)

        Dim grid As DataGrid = Editor.grd_Registers

        grid.CommitEdit(DataGridEditingUnit.Cell, True)

        Dim column As DataGridColumn = grid.ColumnFromDisplayIndex(displayIndex)
        If column Is Nothing Then Exit Sub

        Dim target As New DataGridCellInfo(rowItem, column)
        If Not target.IsValid Then Exit Sub

        grid.CurrentCell = target
        grid.ScrollIntoView(rowItem, column)
        grid.BeginEdit()

    End Sub


    ' ========================================================================
    '  Adding and removing rows
    ' ========================================================================

    ''' <summary>True when the caret is in the rightmost cell of the bottom row.</summary>
    Private Function AtLastCellOfLastRow() As Boolean

        Dim grid As DataGrid = Editor.grd_Registers
        Dim current As DataGridCellInfo = grid.CurrentCell

        If Not current.IsValid OrElse current.Column Is Nothing Then Return False
        If current.Column.DisplayIndex <> grid.Columns.Count - 1 Then Return False
        If Rows Is Nothing OrElse Rows.Count = 0 Then Return False

        Return ReferenceEquals(current.Item, Rows(Rows.Count - 1))

    End Function

    ''' <summary>
    ''' Adds a blank row at the bottom, carrying on the numbering from the row above
    ''' it, and lands on Register Name so typing can continue straight away.
    '''
    ''' Refused for a device whose address is -1: that means it has no register
    ''' addressing at all, so one row is all it can ever have.
    ''' </summary>
    Private Sub AddRegisterRow()

        If Rows Is Nothing Then Exit Sub

        If HasUnaddressedRow() Then
            Warn("This device has a register address of -1, which means it is not " &
                 "register addressed at all." & vbCrLf & vbCrLf &
                 "A device like that has just the one row. Give the row a real address " &
                 "from 0 to 255 if you want more.")
            Exit Sub
        End If

        Dim added As New DeviceRegisterRow With {.RegisterAddress = NextRegisterAddress()}
        Rows.Add(added)

        ' The row's container does not exist until the grid has laid out again.
        Editor.Dispatcher.BeginInvoke(DispatcherPriority.Background,
                                      New Action(Sub() FocusCell(added, RegisterNameColumn)))

    End Sub

    ''' <summary>True when any row carries the "not register addressed" marker.</summary>
    Private Function HasUnaddressedRow() As Boolean

        For Each row As DeviceRegisterRow In Rows
            If If(row.RegisterAddress, String.Empty).Trim() = "-1" Then Return True
        Next

        Return False

    End Function

    ''' <summary>
    ''' One past the address of the bottom row. Blank when that row has no usable
    ''' address, or when carrying on would run off the end of the range - the user
    ''' fills it in by hand from there.
    ''' </summary>
    Private Function NextRegisterAddress() As String

        If Rows.Count = 0 Then Return String.Empty

        Dim previous As Integer
        If Not TryReadAddress(Rows(Rows.Count - 1), previous) Then Return String.Empty

        If previous < 0 OrElse previous >= MaxRegisterAddress Then Return String.Empty

        Return (previous + 1).ToString(CultureInfo.InvariantCulture)

    End Function

    ''' <summary>
    ''' Removes the row the caret is on. The grid always keeps one row, so the last
    ''' one is left alone - blank it out instead, blank rows are not saved.
    ''' </summary>
    Private Sub RemoveCurrentRow()

        If Rows Is Nothing OrElse Rows.Count <= 1 Then Exit Sub

        Dim grid As DataGrid = Editor.grd_Registers
        Dim row As DeviceRegisterRow = TryCast(grid.CurrentCell.Item, DeviceRegisterRow)
        If row Is Nothing Then Exit Sub

        Dim index As Integer = Rows.IndexOf(row)
        If index < 0 Then Exit Sub

        ' An open editor on the row being removed would throw.
        grid.CancelEdit(DataGridEditingUnit.Cell)
        grid.CancelEdit(DataGridEditingUnit.Row)

        Rows.Remove(row)

        Dim landOn As Integer = Math.Min(index, Rows.Count - 1)
        Dim target As DeviceRegisterRow = Rows(landOn)

        Editor.Dispatcher.BeginInvoke(DispatcherPriority.Background,
                                      New Action(Sub() FocusCell(target, 0)))

    End Sub


    ' ========================================================================
    '  Sorting
    '
    '  Only the Register Address column sorts, and it sorts as numbers - the
    '  built-in sort would put "10" in front of "9" because the cell holds text.
    ' ========================================================================

    Private Sub Registers_Sorting(sender As Object, e As DataGridSortingEventArgs)

        ' Always ours, never the grid's own string sort.
        e.Handled = True

        If Rows Is Nothing OrElse Editor.grd_Registers.Columns.Count = 0 Then Exit Sub
        If Not ReferenceEquals(e.Column, Editor.grd_Registers.Columns(0)) Then Exit Sub

        ' A second click on the header turns the order round. An unsorted column has
        ' no direction at all, so the first click sorts ascending.
        Dim descending As Boolean =
            e.Column.SortDirection.HasValue AndAlso
            e.Column.SortDirection.Value = ListSortDirection.Ascending

        SortByRegisterAddress(descending)

        e.Column.SortDirection = If(descending, ListSortDirection.Descending, ListSortDirection.Ascending)

    End Sub

    ''' <summary>
    ''' Puts the rows in register address order. Rows whose address is blank or not a
    ''' number stay at the bottom either way, since there is nothing to sort them by.
    ''' </summary>
    Private Sub SortByRegisterAddress(descending As Boolean)

        If Rows Is Nothing OrElse Rows.Count < 2 Then Exit Sub

        Editor.grd_Registers.CommitEdit(DataGridEditingUnit.Cell, True)
        Editor.grd_Registers.CommitEdit(DataGridEditingUnit.Row, True)

        Dim numbered As New List(Of DeviceRegisterRow)
        Dim unnumbered As New List(Of DeviceRegisterRow)
        Dim address As Integer

        For Each row As DeviceRegisterRow In Rows
            If TryReadAddress(row, address) Then numbered.Add(row) Else unnumbered.Add(row)
        Next

        ' OrderBy is stable, so rows sharing an address keep the order they were in.
        Dim ordered As List(Of DeviceRegisterRow)

        If descending Then
            ordered = numbered.OrderByDescending(Function(row) AddressKey(row)).ToList()
        Else
            ordered = numbered.OrderBy(Function(row) AddressKey(row)).ToList()
        End If

        ordered.AddRange(unnumbered)

        ' Move rather than clear-and-refill, so the rows keep their change handlers
        ' and an already sorted table reports no change at all.
        For target As Integer = 0 To ordered.Count - 1

            Dim source As Integer = Rows.IndexOf(ordered(target))
            If source <> target Then Rows.Move(source, target)

        Next

    End Sub

    ''' <summary>Reads a row's register address, or False if it is not a number.</summary>
    Private Function TryReadAddress(row As DeviceRegisterRow, ByRef value As Integer) As Boolean

        value = 0
        If row Is Nothing Then Return False

        Return Integer.TryParse(If(row.RegisterAddress, String.Empty).Trim(),
                                NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, value)

    End Function

    ''' <summary>Sort key for a row already known to hold a number.</summary>
    Private Function AddressKey(row As DeviceRegisterRow) As Integer

        Dim value As Integer
        TryReadAddress(row, value)

        Return value

    End Function

    ''' <summary>
    ''' The space bar does not come through PreviewTextInput on a TextBox, so it has
    ''' to be turned away at the key. Without this a space slips past the character
    ''' guard and only gets caught later, on Save.
    ''' </summary>
    Private Sub NoSpace_PreviewKeyDown(sender As Object, e As KeyEventArgs)

        If e.Key = Key.Space Then e.Handled = True

    End Sub

    Private Sub DeviceName_PreviewTextInput(sender As Object, e As TextCompositionEventArgs)

        e.Handled = Not IsAlphanumeric(e.Text)

    End Sub

    Private Sub DeviceName_Pasting(sender As Object, e As DataObjectPastingEventArgs)

        Dim pasted As String = TryCast(e.DataObject.GetData(GetType(String)), String)

        If pasted Is Nothing OrElse Not IsAlphanumeric(pasted) Then e.CancelCommand()

    End Sub

    Private Sub Digits_PreviewTextInput(sender As Object, e As TextCompositionEventArgs)

        e.Handled = Not IsDigits(e.Text)

    End Sub

    Private Sub Digits_Pasting(sender As Object, e As DataObjectPastingEventArgs)

        Dim pasted As String = TryCast(e.DataObject.GetData(GetType(String)), String)

        If pasted Is Nothing OrElse Not IsDigits(pasted) Then e.CancelCommand()

    End Sub

    ''' <summary>Letters and digits only - A to Z and 0 to 9, nothing else.</summary>
    Private Function IsAlphanumeric(text As String) As Boolean

        If String.IsNullOrEmpty(text) Then Return False

        For Each character As Char In text
            If Not IsAlphanumeric(character) Then Return False
        Next

        Return True

    End Function

    Private Function IsAlphanumeric(character As Char) As Boolean

        If character >= "A"c AndAlso character <= "Z"c Then Return True
        If character >= "a"c AndAlso character <= "z"c Then Return True
        If character >= "0"c AndAlso character <= "9"c Then Return True

        Return False

    End Function

    Private Function IsDigits(text As String) As Boolean

        If String.IsNullOrEmpty(text) Then Return False

        For Each character As Char In text
            If character < "0"c OrElse character > "9"c Then Return False
        Next

        Return True

    End Function

    ''' <summary>
    ''' The window is closing. Anything that is not a deliberate Save or Cancel -
    ''' the caption X, Alt+F4 - is treated as a Cancel and asks first.
    ''' </summary>
    Public Sub HandleWindowClosing(e As CancelEventArgs)

        If AllowClose Then Exit Sub

        If Not ConfirmAbandon() Then e.Cancel = True

    End Sub


    ' ========================================================================
    '  Loading
    ' ========================================================================

    ''' <summary>Reads a .DEV file into the form. False if it could not be read.</summary>
    Private Function LoadIntoForm(devicefile As String) As Boolean

        If String.IsNullOrWhiteSpace(devicefile) OrElse Not File.Exists(devicefile) Then
            Warn("This device file could not be found:" & vbCrLf & vbCrLf & devicefile)
            Return False
        End If

        Try
            Dim data As DeviceFileData = DeviceLibrary.LoadDeviceFile(devicefile)

            ' A file whose DeviceName is empty falls back to its own file name.
            Dim deviceName As String = If(data.DeviceName, String.Empty).Trim().ToUpperInvariant()
            If deviceName.Length = 0 Then
                deviceName = Path.GetFileNameWithoutExtension(devicefile).Trim().ToUpperInvariant()
            End If

            Editor.txt_DeviceName.Text = deviceName
            Editor.txt_DeviceDescription.Text = If(data.DeviceDescription, String.Empty)
            Editor.txt_BaseAddress.Text = If(data.BaseAddress, String.Empty)
            Editor.txt_AddressBits.Text = If(data.AddressBits, String.Empty)

            OriginalDeviceName = deviceName

            Rows.Clear()

            If data.Registers IsNot Nothing Then
                For Each register As DeviceRegisterData In data.Registers
                    If register Is Nothing Then Continue For
                    Rows.Add(New DeviceRegisterRow With {
                        .RegisterAddress = register.RegisterAddress.ToString(CultureInfo.InvariantCulture),
                        .RegisterName = If(register.RegisterName, String.Empty),
                        .D7 = If(register.D7, String.Empty),
                        .D6 = If(register.D6, String.Empty),
                        .D5 = If(register.D5, String.Empty),
                        .D4 = If(register.D4, String.Empty),
                        .D3 = If(register.D3, String.Empty),
                        .D2 = If(register.D2, String.Empty),
                        .D1 = If(register.D1, String.Empty),
                        .D0 = If(register.D0, String.Empty)
                    })
                Next
            End If

            Return True

        Catch ex As Exception
            Warn("This device file could not be read:" & vbCrLf & vbCrLf & ex.Message)
            Return False
        End Try

    End Function


    ' ========================================================================
    '  Device name field appearance
    ' ========================================================================

    ''' <summary>Edit mode: the name is the file name, so it cannot be changed.</summary>
    Private Sub LockDeviceName()

        Editor.txt_DeviceName.IsReadOnly = True
        Editor.txt_DeviceName.Background = TryCast(Editor.TryFindResource("Brush_Field_Locked_Background"), Brush)
        Editor.txt_DeviceName.Foreground = TryCast(Editor.TryFindResource("Brush_Field_Locked_Text"), Brush)
        Editor.txt_DeviceName.ToolTip = "The device name cannot be changed here. Use Clone Device to make a copy under a new name."

        Editor.txt_DeviceDescription.Focus()

    End Sub

    ''' <summary>Clone mode: light red until the name is actually different.</summary>
    Private Sub FlagDeviceNameForRename()

        Editor.txt_DeviceName.Background = TryCast(Editor.TryFindResource("Brush_Field_Rename_Background"), Brush)
        Editor.txt_DeviceName.Foreground = TryCast(Editor.TryFindResource("Brush_Field_Rename_Text"), Brush)
        Editor.txt_DeviceName.CaretBrush = TryCast(Editor.TryFindResource("Brush_Field_Rename_Text"), Brush)
        Editor.txt_DeviceName.ToolTip = "This is still the original name. Give the clone a new one."

        Editor.txt_DeviceName.Focus()
        Editor.txt_DeviceName.SelectAll()

    End Sub

    ''' <summary>
    ''' Back to the normal light-blue-on-85%-black. ClearValue drops the local
    ''' colours so the field falls back to Style_Editor_Field.
    ''' </summary>
    Private Sub ResetDeviceNameColours()

        Editor.txt_DeviceName.ClearValue(TextBox.BackgroundProperty)
        Editor.txt_DeviceName.ClearValue(TextBox.ForegroundProperty)
        Editor.txt_DeviceName.ClearValue(TextBox.CaretBrushProperty)
        Editor.txt_DeviceName.ToolTip = Nothing

    End Sub

    ''' <summary>
    ''' Runs on every keystroke in the device name: forces upper case, then - in
    ''' clone mode only - keeps the field flagged until the name is genuinely
    ''' different from the one it was cloned from.
    ''' </summary>
    Public Sub DeviceNameEdited()

        If Editor Is Nothing Then Exit Sub
        If SuppressNameSync Then Exit Sub

        MarkDirty()
        ForceDeviceNameUpperCase()

        If Editor.EditorMode <> DeviceEditorMode.mode_Clone Then Exit Sub

        Dim current As String = Editor.txt_DeviceName.Text.Trim()

        If current.Length = 0 OrElse
           String.Equals(current, OriginalDeviceName, StringComparison.OrdinalIgnoreCase) Then
            FlagDeviceNameForRenameQuietly()
        Else
            ResetDeviceNameColours()
        End If

    End Sub

    ''' <summary>
    ''' Rewrites the device name in upper case as it is typed, putting the caret
    ''' back where it was. The guard flag keeps the TextChanged handler from
    ''' re-entering when the text is replaced.
    ''' </summary>
    Private Sub ForceDeviceNameUpperCase()

        Dim typed As String = Editor.txt_DeviceName.Text
        Dim upper As String = typed.ToUpperInvariant()

        If String.Equals(typed, upper, StringComparison.Ordinal) Then Exit Sub

        Dim caret As Integer = Editor.txt_DeviceName.CaretIndex

        SuppressNameSync = True
        Editor.txt_DeviceName.Text = upper
        SuppressNameSync = False

        Editor.txt_DeviceName.CaretIndex = caret

    End Sub

    ''' <summary>As FlagDeviceNameForRename, without stealing focus or selecting.</summary>
    Private Sub FlagDeviceNameForRenameQuietly()

        Editor.txt_DeviceName.Background = TryCast(Editor.TryFindResource("Brush_Field_Rename_Background"), Brush)
        Editor.txt_DeviceName.Foreground = TryCast(Editor.TryFindResource("Brush_Field_Rename_Text"), Brush)
        Editor.txt_DeviceName.CaretBrush = TryCast(Editor.TryFindResource("Brush_Field_Rename_Text"), Brush)
        Editor.txt_DeviceName.ToolTip = "This is still the original name. Give the clone a new one."

    End Sub


    ' ========================================================================
    '  Base Address readout
    '
    '  The arithmetic lives in Modules\I2CAddress.vb. This only decides what to
    '  put in the label and what colour to put it in.
    ' ========================================================================

    ''' <summary>Base Address was typed in.</summary>
    Public Sub BaseAddressEdited()

        MarkDirty()
        RefreshAddressReadout()

    End Sub

    ''' <summary>Address bits was typed in - it decides how much of the readout is masked.</summary>
    Public Sub AddressBitsEdited()

        MarkDirty()
        RefreshAddressReadout()

    End Sub

    ''' <summary>
    ''' Rebuilds the readout from both boxes. An unusable Base Address is reported
    ''' first - there is nothing to mask until the address itself makes sense.
    ''' </summary>
    Private Sub RefreshAddressReadout()

        If Editor Is Nothing Then Exit Sub

        Dim address As Integer
        Dim addressState As I2CAddressState = I2CAddress.Check(Editor.txt_BaseAddress.Text, address)

        Select Case addressState

            Case I2CAddressState.state_Empty
                ShowReadout(String.Empty, True, Nothing)
                Exit Sub

            Case I2CAddressState.state_OutOfRange
                ShowReadout("0 - 255 only", False, "Base Address must be 0 to 255.")
                Exit Sub

            Case I2CAddressState.state_OddEightBit
                ShowReadout("must be even", False,
                            "Above 127 the value is an 8 bit address, so its bottom bit is the " &
                            "read/write flag and has to be zero.")
                Exit Sub

        End Select

        Dim bits As Integer
        Dim bitsState As I2CAddressState = I2CAddress.CheckAddressBits(Editor.txt_AddressBits.Text, bits)

        If bitsState = I2CAddressState.state_OutOfRange Then
            ShowReadout("bits 0 - 7 only", False, "Address bits must be 0 to 7.")
            Exit Sub
        End If

        ' A blank Address bits box simply masks nothing.
        Dim sevenBit As Integer = I2CAddress.SevenBitAddress(address)

        Dim hint As String =
            "Base Address " & address.ToString(CultureInfo.InvariantCulture) &
            " means 7 bit address " & sevenBit.ToString(CultureInfo.InvariantCulture) &
            " (0x" & sevenBit.ToString("X2", CultureInfo.InvariantCulture) & ")"

        If bits > 0 Then
            hint &= vbCrLf & "X marks the " & bits.ToString(CultureInfo.InvariantCulture) &
                    " bit(s) the address pins choose."
        End If

        ShowReadout(I2CAddress.FormatSevenBit(sevenBit, bits), True, hint)

    End Sub

    Private Sub ShowReadout(text As String, good As Boolean, hint As String)

        Editor.lbl_AddressReadout.Text = text
        Editor.lbl_AddressReadout.Foreground =
            TryCast(Editor.TryFindResource(If(good, "Brush_Address_Good", "Brush_Address_Bad")), Brush)
        Editor.lbl_AddressReadout.ToolTip = hint

    End Sub


    ' ========================================================================
    '  Saving
    ' ========================================================================

    ''' <summary>
    ''' Checks the form over and writes the .DEV file. Every failure leaves the form
    ''' open and unchanged so the user can fix whatever was wrong.
    ''' </summary>
    Public Sub SaveDevice()

        If Editor Is Nothing Then Exit Sub

        ' Commit a cell that is still being typed into, or its text would be missed.
        Editor.grd_Registers.CommitEdit(DataGridEditingUnit.Cell, True)
        Editor.grd_Registers.CommitEdit(DataGridEditingUnit.Row, True)

        ' The file is always written in register address order, whatever order the
        ' table happens to be showing.
        SortByRegisterAddress(False)
        If Editor.grd_Registers.Columns.Count > 0 Then
            Editor.grd_Registers.Columns(0).SortDirection = ListSortDirection.Ascending
        End If

        ' Trimmed and upper cased, which is the only form a device name is stored in.
        Dim deviceName As String = Editor.txt_DeviceName.Text.Trim().ToUpperInvariant()

        If deviceName.Length = 0 Then
            Warn("Device Name is empty. The file is named after it, so it has to be filled in.")
            Editor.txt_DeviceName.Focus()
            Exit Sub
        End If

        If Not DeviceLibrary.IsUsableDeviceName(deviceName) Then
            Warn("Device Name may only contain letters and numbers." & vbCrLf & vbCrLf &
                 "Nothing else is allowed - no spaces, dashes or punctuation.")
            Editor.txt_DeviceName.Focus()
            Editor.txt_DeviceName.SelectAll()
            Exit Sub
        End If

        ' Put the cleaned-up name back, so what gets saved is what is on screen.
        If Not String.Equals(Editor.txt_DeviceName.Text, deviceName, StringComparison.Ordinal) Then
            SuppressNameSync = True
            Editor.txt_DeviceName.Text = deviceName
            SuppressNameSync = False
        End If

        Dim baseAddressValue As Integer
        Dim baseAddressState As I2CAddressState = I2CAddress.Check(Editor.txt_BaseAddress.Text, baseAddressValue)

        Select Case baseAddressState

            Case I2CAddressState.state_Empty
                Warn("Base Address is empty. It has to be a value from 0 to 255.")
                Editor.txt_BaseAddress.Focus()
                Exit Sub

            Case I2CAddressState.state_OutOfRange
                Warn("Base Address must be a value from 0 to 255.")
                Editor.txt_BaseAddress.Focus()
                Editor.txt_BaseAddress.SelectAll()
                Exit Sub

            Case I2CAddressState.state_OddEightBit
                Warn("Base Address " & baseAddressValue.ToString(CultureInfo.InvariantCulture) &
                     " is odd." & vbCrLf & vbCrLf &
                     "Anything above 127 is read as an 8 bit address, where the bottom bit is the " &
                     "read/write flag and has to be zero. Use " &
                     (baseAddressValue And &HFE).ToString(CultureInfo.InvariantCulture) &
                     ", or enter the 7 bit address " &
                     I2CAddress.SevenBitAddress(baseAddressValue).ToString(CultureInfo.InvariantCulture) & " instead.")
                Editor.txt_BaseAddress.Focus()
                Editor.txt_BaseAddress.SelectAll()
                Exit Sub

        End Select

        Dim addressBitsValue As Integer
        Dim addressBitsState As I2CAddressState =
            I2CAddress.CheckAddressBits(Editor.txt_AddressBits.Text, addressBitsValue)

        If addressBitsState <> I2CAddressState.state_Good Then
            Warn("Address bits must be a value from 0 to " &
                 I2CAddress.MaxAddressBits.ToString(CultureInfo.InvariantCulture) & "." & vbCrLf & vbCrLf &
                 "That is how many of the address bits the part's address pins choose. " &
                 "Use 0 for a part whose address is fixed.")
            Editor.txt_AddressBits.Focus()
            Editor.txt_AddressBits.SelectAll()
            Exit Sub
        End If

        Dim registers As List(Of DeviceRegisterData) = CollectRegisters()
        If registers Is Nothing Then Exit Sub          ' CollectRegisters has already explained

        ' Edit mode writes back the very file it opened. Every other mode derives
        ' the path from the name, which is what makes the overwrite question real.
        Dim targetPath As String

        If Editor.EditorMode = DeviceEditorMode.mode_Edit Then
            targetPath = Editor.DeviceFilePath
        Else
            targetPath = DeviceLibrary.PathFor(deviceName)
        End If

        Select Case Editor.EditorMode

            Case DeviceEditorMode.mode_Clone
                If File.Exists(targetPath) Then
                    Warn("There is already a device called """ & deviceName & """." & vbCrLf & vbCrLf &
                         "A clone needs a name of its own. Change the Device Name, or cancel the form.")
                    Editor.txt_DeviceName.Focus()
                    Editor.txt_DeviceName.SelectAll()
                    Exit Sub
                End If

            Case DeviceEditorMode.mode_New
                If File.Exists(targetPath) Then
                    Dim answer As MessageBoxResult = MessageBox.Show(
                        Editor,
                        """" & Path.GetFileName(targetPath) & """ already exists." & vbCrLf & vbCrLf &
                        "Overwrite it?",
                        DialogTitle, MessageBoxButton.YesNo, MessageBoxImage.Question)

                    ' No means back to the form, so the name can be changed.
                    If answer <> MessageBoxResult.Yes Then
                        Editor.txt_DeviceName.Focus()
                        Editor.txt_DeviceName.SelectAll()
                        Exit Sub
                    End If
                End If

        End Select

        Try
            Dim data As New DeviceFileData With {
                .DeviceName = deviceName,
                .DeviceDescription = Editor.txt_DeviceDescription.Text.Trim(),
                .BaseAddress = Editor.txt_BaseAddress.Text.Trim(),
                .AddressBits = Editor.txt_AddressBits.Text.Trim(),
                .Registers = registers
            }

            DeviceLibrary.SaveDeviceFile(targetPath, data)

        Catch ex As Exception
            Warn("The device could not be saved:" & vbCrLf & vbCrLf & ex.Message)
            Exit Sub
        End Try

        AppCore.SetStatus("Saved " & Path.GetFileName(targetPath) &
                          " (" & registers.Count.ToString(CultureInfo.InvariantCulture) & " registers)")

        AllowClose = True
        Editor.DialogResult = True
        Editor.Close()

    End Sub

    ''' <summary>
    ''' Turns the grid into register records, applying every rule:
    '''
    '''   - blank rows are dropped
    '''   - a blank address on a row that has other data becomes -1
    '''   - addresses must be -1, or 0 to 255
    '''   - addresses must not repeat
    '''   - if any row is -1 it has to be the only row
    '''
    ''' Returns Nothing after telling the user what is wrong.
    ''' </summary>
    Private Function CollectRegisters() As List(Of DeviceRegisterData)

        Dim registers As New List(Of DeviceRegisterData)
        Dim rowOfAddress As New Dictionary(Of Integer, Integer)
        Dim unaddressedRows As Integer = 0
        Dim rowNumber As Integer = 0

        For Each row As DeviceRegisterRow In Rows

            rowNumber += 1

            If row Is Nothing OrElse row.IsBlank() Then Continue For

            Dim addressText As String = If(row.RegisterAddress, String.Empty).Trim()
            Dim address As Integer

            If addressText.Length = 0 Then
                ' Data but no address: the device is not register addressed.
                address = -1
                unaddressedRows += 1
            Else
                If Not Integer.TryParse(addressText, NumberStyles.AllowLeadingSign,
                                        CultureInfo.InvariantCulture, address) Then
                    Warn("Row " & rowNumber.ToString(CultureInfo.InvariantCulture) & ": """ & addressText &
                         """ is not a number." & vbCrLf & vbCrLf &
                         "Register Address must be 0 to 255, or -1.")
                    Return Nothing
                End If

                If address < MinRegisterAddress OrElse address > MaxRegisterAddress Then
                    Warn("Row " & rowNumber.ToString(CultureInfo.InvariantCulture) & ": " &
                         address.ToString(CultureInfo.InvariantCulture) & " is out of range." & vbCrLf & vbCrLf &
                         "Register Address must be 0 to 255, or -1.")
                    Return Nothing
                End If

                If address = -1 Then unaddressedRows += 1
            End If

            If rowOfAddress.ContainsKey(address) Then
                Warn("Duplicate Register Address " & address.ToString(CultureInfo.InvariantCulture) & "." & vbCrLf & vbCrLf &
                     "It is used by row " & rowOfAddress(address).ToString(CultureInfo.InvariantCulture) &
                     " and row " & rowNumber.ToString(CultureInfo.InvariantCulture) & "." & vbCrLf & vbCrLf &
                     "Every register needs its own address.")
                Return Nothing
            End If

            rowOfAddress.Add(address, rowNumber)

            registers.Add(New DeviceRegisterData With {
                .RegisterAddress = address,
                .RegisterName = If(row.RegisterName, String.Empty).Trim(),
                .D7 = If(row.D7, String.Empty).Trim(),
                .D6 = If(row.D6, String.Empty).Trim(),
                .D5 = If(row.D5, String.Empty).Trim(),
                .D4 = If(row.D4, String.Empty).Trim(),
                .D3 = If(row.D3, String.Empty).Trim(),
                .D2 = If(row.D2, String.Empty).Trim(),
                .D1 = If(row.D1, String.Empty).Trim(),
                .D0 = If(row.D0, String.Empty).Trim()
            })

        Next

        If unaddressedRows > 0 AndAlso registers.Count > 1 Then
            Warn("A register with no address means the device is not register addressed," & vbCrLf &
                 "so it can only have the one row." & vbCrLf & vbCrLf &
                 "There are " & registers.Count.ToString(CultureInfo.InvariantCulture) &
                 " rows of data. Either give every row an address, or leave just one row.")
            Return Nothing
        End If

        Return registers

    End Function


    ' ========================================================================
    '  Cancelling
    ' ========================================================================

    Public Sub CancelEditor()

        If Editor Is Nothing Then Exit Sub

        If Not ConfirmAbandon() Then Exit Sub

        AllowClose = True
        Editor.DialogResult = False
        Editor.Close()

    End Sub

    ''' <summary>
    ''' Nothing typed in means nothing to lose, so an untouched form just closes.
    ''' </summary>
    Private Function ConfirmAbandon() As Boolean

        If Not IsDirty Then Return True

        Dim answer As MessageBoxResult = MessageBox.Show(
            Editor,
            "Close the Device Editor without saving?" & vbCrLf & vbCrLf &
            "Anything typed in will be lost.",
            DialogTitle, MessageBoxButton.YesNo, MessageBoxImage.Question)

        Return answer = MessageBoxResult.Yes

    End Function


    ' ========================================================================
    '  Small helpers
    ' ========================================================================

    Private Sub Warn(message As String)

        MessageBox.Show(Editor, message, DialogTitle, MessageBoxButton.OK, MessageBoxImage.Warning)

    End Sub

End Module
