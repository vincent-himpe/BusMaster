' ============================================================================
'  Modules\EventLogCore.vb
'
'  All of the Event Log's active code. The window's handlers only call in here.
'
'  Typing in the marker column is the whole point of the grid: one character,
'  upper cased, and the caret drops to the next line ready for the next one. A
'  label ("L") also pushes a fresh line in below itself.
'
'  Logs are written as plain CSV so they can be read anywhere. Nothing about the
'  colours is stored - they are derived from the marker column, so a log read back
'  looks exactly as it did.
' ============================================================================

Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.ComponentModel
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Windows.Threading
Imports Microsoft.VisualBasic

Public Module EventLogCore

    Public Const LogExtension As String = ".buslog"
    Public Const LogFolderName As String = "buslogs"

    Private Const WindowTitle As String = "Event Log"

    ''' <summary>Display index of the marker column - where typing lands.</summary>
    Private Const MarkerColumn As Integer = 1

    ''' <summary>
    ''' Display index of the line number column. Read-only, and the only place a
    ''' double click sets a breakpoint.
    ''' </summary>
    Private Const NumberColumn As Integer = 0

    ''' <summary>
    ''' Leftmost column the caret may sit in. Column 0 is the line number, which is
    ''' read-only, so arrowing left stops at the marker.
    ''' </summary>
    Private Const FirstEditableColumn As Integer = 1

    Private Log As EventLogWindow
    Private Rows As ObservableCollection(Of EventLogRow)

    ''' <summary>File the log was last saved to or loaded from, or an empty string.</summary>
    Private LogFilePath As String = String.Empty

    ''' <summary>Guards the renumbering pass against re-entering itself.</summary>
    Private Renumbering As Boolean = False


    ' ========================================================================
    '  Decimal / hexadecimal
    '
    '  The log's own switch, with nothing to do with the main window's. The two
    '  are read for different reasons and at different times - the registers on
    '  screen are what a part holds now, the log is what went past - and having to
    '  change one to read the other would be a nuisance rather than a convenience.
    '
    '  Host, Reg and Value are stored in decimal whatever this says, and that is
    '  what is written to the file. Only the three columns on screen change, and
    '  the comment column never does: it is prose.
    ' ========================================================================

    ''' <summary>
    ''' Whether the log's three number columns are being read as hexadecimal. Public
    ''' because the row itself converts through it - see EventLogRow.HostText.
    ''' </summary>
    Public Property ShowingHexadecimal As Boolean = False

    Public Sub ToggleRadix()

        ShowingHexadecimal = Not ShowingHexadecimal

        ' Nothing has moved; the columns are just read a different way now.
        If Rows IsNot Nothing Then
            For Each row As EventLogRow In Rows
                row.AnnounceRadix()
            Next
        End If

        RefreshRadixButton()

        AppCore.SetStatus("Event log values shown in " &
                          If(ShowingHexadecimal, "hexadecimal", "decimal"))

    End Sub

    ''' <summary>Keeps the toolbar button showing which way the log is being read.</summary>
    Private Sub RefreshRadixButton()

        If Log Is Nothing Then Exit Sub

        Log.tlbr_Log_ValueType.Content = Radix.Caption(ShowingHexadecimal)
        Log.tlbr_Log_ValueType.Foreground = TryCast(
            Log.TryFindResource(Radix.BrushKey(ShowingHexadecimal)), Brush)

    End Sub


    ' ========================================================================
    '  Where logs live
    ' ========================================================================

    Public ReadOnly Property LogFolder As String
        Get
            Return Path.Combine(AppContext.BaseDirectory, LogFolderName)
        End Get
    End Property

    Public Sub EnsureLogFolder()

        If Not Directory.Exists(LogFolder) Then Directory.CreateDirectory(LogFolder)

    End Sub


    ' ========================================================================
    '  Start-up and shutdown
    ' ========================================================================

    Public Sub Attach(window As EventLogWindow)

        Log = window

        WindowTheme.ApplyDarkTitleBar(Log)

        Rows = New ObservableCollection(Of EventLogRow)
        AddHandler Rows.CollectionChanged, AddressOf Rows_CollectionChanged

        Log.grd_Log.ItemsSource = Rows
        AddHandler Log.grd_Log.PreviewKeyDown, AddressOf Log_PreviewKeyDown
        AddHandler Log.grd_Log.PreviewTextInput, AddressOf Log_PreviewTextInput
        AddHandler Log.grd_Log.MouseDoubleClick, AddressOf Log_MouseDoubleClick

        ' There has to be somewhere to start typing.
        Rows.Add(New EventLogRow)

        RefreshRadixButton()
        RefreshTitle()

    End Sub

    ''' <summary>
    ''' The log is put away, not thrown away: it is created once and hidden, so
    ''' whatever has been typed is still there next time it is opened.
    ''' </summary>
    Public Sub HandleWindowClosing(e As CancelEventArgs)

        If Log Is Nothing Then Exit Sub

        e.Cancel = True
        Log.Hide()

    End Sub

    ''' <summary>Brings the log up, creating nothing - it already exists.</summary>
    Public Sub ShowLog()

        If Log Is Nothing Then Exit Sub

        Log.Show()

        If Log.WindowState = WindowState.Minimized Then Log.WindowState = WindowState.Normal

        Log.Activate()

    End Sub

    Private Sub RefreshTitle()

        If Log Is Nothing Then Exit Sub

        If LogFilePath.Length = 0 Then
            Log.Title = WindowTitle
        Else
            Log.Title = WindowTitle & " - " & Path.GetFileName(LogFilePath)
        End If

    End Sub


    ' ========================================================================
    '  Row numbering
    ' ========================================================================

    Private Sub Rows_CollectionChanged(sender As Object, e As NotifyCollectionChangedEventArgs)

        Renumber()

    End Sub

    ''' <summary>Puts the line numbers back in sequence after rows come or go.</summary>
    Private Sub Renumber()

        If Renumbering OrElse Rows Is Nothing Then Exit Sub

        Renumbering = True

        For index As Integer = 0 To Rows.Count - 1
            Rows(index).Number = index + 1
        Next

        Renumbering = False

    End Sub


    ' ========================================================================
    '  Typing
    ' ========================================================================

    ''' <summary>
    ''' Only the listed markers and a space get into the marker column, and only one
    ''' character of it.
    ''' </summary>
    Private Sub Log_PreviewTextInput(sender As Object, e As TextCompositionEventArgs)

        If CurrentColumnIndex() <> MarkerColumn Then Exit Sub

        Dim typed As String = If(e.Text, String.Empty).ToUpperInvariant()

        If typed.Length <> 1 Then
            e.Handled = True
            Exit Sub
        End If

        ' A space clears the cell; anything else has to be a known marker.
        If typed <> " " AndAlso Not EventLogMarkers.Allowed.Contains(typed) Then
            e.Handled = True
            Exit Sub
        End If

        Dim index As Integer = CurrentRowIndex()
        If index < 0 Then
            e.Handled = True
            Exit Sub
        End If

        e.Handled = True

        If typed = EventLogMarkers.Label Then

            ' A label makes room for itself: everything from the cursor down is
            ' pushed along and the new line that lands here becomes the label.
            InsertRowAt(index)
            Rows(index).Marker = typed

        Else

            ' Written straight to the row rather than let into the editor: the cell
            ' holds one character, and typing over an existing one has to replace it.
            Rows(index).Marker = If(typed = " ", String.Empty, typed)

        End If

        MoveToRow(index + 1)

    End Sub

    Private Sub Log_PreviewKeyDown(sender As Object, e As KeyEventArgs)

        Dim control As Boolean = (Keyboard.Modifiers And ModifierKeys.Control) = ModifierKeys.Control
        Dim editing As TextBox = TryCast(Keyboard.FocusedElement, TextBox)

        Select Case e.Key

            Case Key.Up
                ' Vertical arrows end the entry rather than move within the text.
                MoveToRow(CurrentRowIndex() - 1, False)
                e.Handled = True

            Case Key.Down
                MoveToRow(CurrentRowIndex() + 1, False)
                e.Handled = True

            Case Key.Left
                ' Only steps out of the cell once the caret is already at the front.
                If editing IsNot Nothing AndAlso CaretAtStart(editing) Then
                    MoveToColumn(CurrentColumnIndex() - 1)
                    e.Handled = True
                End If

            Case Key.Right
                If editing IsNot Nothing AndAlso CaretAtEnd(editing) Then
                    MoveToColumn(CurrentColumnIndex() + 1)
                    e.Handled = True
                End If

            Case Key.Insert
                ' The new line takes the cursor's place, and the cursor goes with it -
                ' otherwise the caret ends up on the row that was pushed down, which
                ' is not the one the user just made room for.
                Dim at As Integer = CurrentRowIndex()
                InsertRowAt(at)
                FocusCell(at, MarkerColumn)
                e.Handled = True

            Case Key.Delete
                If control Then
                    RemoveCurrentRow()
                    e.Handled = True
                End If

            Case Key.Space
                ' Space never reaches PreviewTextInput on a grid cell, so it is
                ' handled here: clear the marker and drop down a line.
                If CurrentColumnIndex() = MarkerColumn Then
                    Dim index As Integer = CurrentRowIndex()
                    If index >= 0 Then
                        Rows(index).Marker = String.Empty
                        e.Handled = True
                        MoveToRow(index + 1)
                    End If
                End If

        End Select

    End Sub

    ''' <summary>
    ''' Steps to another line, keeping the column. Typing a marker grows the log off
    ''' the bottom so a run can be entered without stopping; arrowing about does not,
    ''' it just stops at the last line.
    ''' </summary>
    Private Sub MoveToRow(rowIndex As Integer, Optional growPastTheEnd As Boolean = True)

        If Rows Is Nothing Then Exit Sub

        Dim column As Integer = CurrentColumnIndex()
        If column < FirstEditableColumn Then column = MarkerColumn

        Dim target As Integer = rowIndex

        If target < 0 Then target = 0

        If target > Rows.Count - 1 Then
            If Not growPastTheEnd Then
                target = Rows.Count - 1
            Else
                Rows.Add(New EventLogRow)
                target = Rows.Count - 1
            End If
        End If

        FocusCell(target, column)

    End Sub

    ''' <summary>
    ''' Steps to another column on the same line. Hard stop at both ends - it never
    ''' rolls onto the line above or below.
    ''' </summary>
    Private Sub MoveToColumn(columnIndex As Integer)

        Dim grid As DataGrid = Log.grd_Log

        Dim target As Integer = columnIndex
        If target < FirstEditableColumn Then target = FirstEditableColumn
        If target > grid.Columns.Count - 1 Then target = grid.Columns.Count - 1

        FocusCell(CurrentRowIndex(), target)

    End Sub

    ''' <summary>
    ''' Double clicking a line number turns its breakpoint on or off. Anywhere else
    ''' in the grid is left alone, so a double click still selects a word to edit.
    ''' </summary>
    Private Sub Log_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs)

        Dim cell As DataGridCell = CellUnder(TryCast(e.OriginalSource, DependencyObject))
        If cell Is Nothing OrElse cell.Column Is Nothing Then Exit Sub

        If cell.Column.DisplayIndex <> NumberColumn Then Exit Sub

        Dim row As EventLogRow = TryCast(cell.DataContext, EventLogRow)
        If row Is Nothing Then Exit Sub

        row.IsBreakpoint = Not row.IsBreakpoint
        e.Handled = True

    End Sub

    ''' <summary>The grid cell an element sits in, or Nothing.</summary>
    Private Function CellUnder(source As DependencyObject) As DataGridCell

        Dim node As DependencyObject = source

        While node IsNot Nothing

            Dim cell As DataGridCell = TryCast(node, DataGridCell)
            If cell IsNot Nothing Then Return cell

            node = VisualTreeHelper.GetParent(node)

        End While

        Return Nothing

    End Function

    ''' <summary>Caret sitting in front of the first character, nothing selected.</summary>
    Private Function CaretAtStart(box As TextBox) As Boolean

        Return box.SelectionLength = 0 AndAlso box.CaretIndex = 0

    End Function

    ''' <summary>Caret sitting past the last character, nothing selected.</summary>
    Private Function CaretAtEnd(box As TextBox) As Boolean

        Return box.SelectionLength = 0 AndAlso box.CaretIndex >= box.Text.Length

    End Function


    ' ========================================================================
    '  Rows
    ' ========================================================================

    Public Sub InsertRowAt(index As Integer)

        If Rows Is Nothing Then Exit Sub

        Dim at As Integer = index
        If at < 0 Then at = 0
        If at > Rows.Count Then at = Rows.Count

        Rows.Insert(at, New EventLogRow)

    End Sub

    Public Sub RemoveCurrentRow()

        If Rows Is Nothing OrElse Rows.Count <= 1 Then Exit Sub

        Dim index As Integer = CurrentRowIndex()
        If index < 0 Then Exit Sub

        Log.grd_Log.CancelEdit(DataGridEditingUnit.Cell)
        Log.grd_Log.CancelEdit(DataGridEditingUnit.Row)

        Rows.RemoveAt(index)

        FocusCell(Math.Min(index, Rows.Count - 1), MarkerColumn)

    End Sub

    ''' <summary>
    ''' Puts a line on the end of the log, wherever the caret happens to be. This is
    ''' the way in for anything that records itself rather than being typed.
    '''
    ''' The blank line the log always keeps to type on is filled in rather than
    ''' pushed down, so recorded lines do not end up with a gap above them, and a
    ''' fresh blank one is left behind.
    ''' </summary>
    Public Sub AppendEntry(marker As String, host As String, reg As String,
                           value As String, comment As String)

        If Rows Is Nothing Then Exit Sub

        ' A DataGrid will not let its source collection change while a row is open
        ' for editing, and the log leaves a cell open whenever it has been used.
        ' Committing first keeps whatever was typed and lets the new row through.
        If Log IsNot Nothing Then
            Log.grd_Log.CommitEdit(DataGridEditingUnit.Cell, True)
            Log.grd_Log.CommitEdit(DataGridEditingUnit.Row, True)
        End If

        Dim row As EventLogRow

        If Rows.Count > 0 AndAlso Rows(Rows.Count - 1).IsBlank() Then
            row = Rows(Rows.Count - 1)
        Else
            row = New EventLogRow
            Rows.Add(row)
        End If

        row.Marker = marker
        row.Host = host
        row.Reg = reg
        row.Value = value
        row.Comment = comment

        Rows.Add(New EventLogRow)

        ' Only worth scrolling if anyone is looking.
        If Log IsNot Nothing AndAlso Log.IsVisible Then
            Log.grd_Log.ScrollIntoView(Rows(Rows.Count - 1))
        End If

    End Sub

    ''' <summary>Empties the log and starts again on one blank line.</summary>
    Public Sub ClearLog()

        If Rows Is Nothing Then Exit Sub

        Log.grd_Log.CancelEdit(DataGridEditingUnit.Cell)
        Log.grd_Log.CancelEdit(DataGridEditingUnit.Row)

        Rows.Clear()
        Rows.Add(New EventLogRow)

        FocusCell(0, MarkerColumn)
        AppCore.SetStatus("Event log cleared")

    End Sub


    ' ========================================================================
    '  Grid position
    ' ========================================================================

    Private Function CurrentRow() As EventLogRow

        Return TryCast(Log.grd_Log.CurrentCell.Item, EventLogRow)

    End Function

    Private Function CurrentRowIndex() As Integer

        Dim row As EventLogRow = CurrentRow()
        If row Is Nothing Then Return -1

        Return Rows.IndexOf(row)

    End Function

    Private Function CurrentColumnIndex() As Integer

        Dim current As DataGridCellInfo = Log.grd_Log.CurrentCell

        If Not current.IsValid OrElse current.Column Is Nothing Then Return -1

        Return current.Column.DisplayIndex

    End Function

    Private Sub FocusCell(rowIndex As Integer, columnIndex As Integer)

        If rowIndex < 0 OrElse rowIndex > Rows.Count - 1 Then Exit Sub

        Dim grid As DataGrid = Log.grd_Log
        Dim item As EventLogRow = Rows(rowIndex)

        ' The row container may not exist until the grid has laid out again.
        Log.Dispatcher.BeginInvoke(DispatcherPriority.Background,
            New Action(Sub()
                           Dim column As DataGridColumn = grid.ColumnFromDisplayIndex(columnIndex)
                           If column Is Nothing Then Exit Sub

                           Dim target As New DataGridCellInfo(item, column)
                           If Not target.IsValid Then Exit Sub

                           grid.CurrentCell = target
                           grid.ScrollIntoView(item, column)

                           grid.SelectedCells.Clear()
                           grid.SelectedCells.Add(target)

                           ' Opening the cell is what actually puts keyboard focus
                           ' inside it, which is what the active row and cell
                           ' highlights key off. Typing in the marker column is
                           ' intercepted before the editor sees it.
                           grid.Focus()
                           grid.BeginEdit()
                       End Sub))

    End Sub


    ' ========================================================================
    '  Saving and loading
    '
    '  Plain CSV, the whole table, one line per row.
    ' ========================================================================

    ''' <summary>
    ''' Asks for a name and writes ".\buslogs\<project>-<name>.buslog". The project
    ''' name is the prefix so a log always says which layout produced it.
    ''' </summary>
    Public Sub SaveLogAs()

        Dim answer As String = StringPrompt.Ask(
            Log, WindowTitle,
            "Name for this log:" & vbCrLf & vbCrLf &
            "It will be saved as " & AppCore.ProjectNameForFiles() & "-<name>" & LogExtension)

        If answer Is Nothing Then Exit Sub

        Dim entered As String = answer.Trim()
        If entered.Length = 0 Then Exit Sub

        If entered.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 Then
            MessagePrompt.ShowWarning(Log, WindowTitle,
                                      "That name contains characters that cannot be used in a file name.")
            Exit Sub
        End If

        Dim fileName As String = AppCore.ProjectNameForFiles() & "-" & entered & LogExtension

        WriteLog(Path.Combine(LogFolder, fileName))

    End Sub

    ''' <summary>
    ''' Overwrites the file the log came from, without asking. With no file behind it
    ''' yet, this is a Save As.
    ''' </summary>
    Public Sub SaveLog()

        If LogFilePath.Length = 0 Then
            SaveLogAs()
            Exit Sub
        End If

        WriteLog(LogFilePath)

    End Sub

    Private Sub WriteLog(filePath As String)

        Try
            EnsureLogFolder()

            Dim text As New StringBuilder()
            text.AppendLine("Number,Marker,Host,Reg,Value,Comment")

            For Each row As EventLogRow In Rows
                text.AppendLine(String.Join(",", {
                    row.Number.ToString(CultureInfo.InvariantCulture),
                    Quote(row.Marker),
                    Quote(row.Host),
                    Quote(row.Reg),
                    Quote(row.Value),
                    Quote(row.Comment)}))
            Next

            File.WriteAllText(filePath, text.ToString())

            LogFilePath = filePath
            RefreshTitle()
            AppCore.SetStatus("Saved " & Path.GetFileName(filePath))

        Catch ex As Exception
            MessagePrompt.ShowError(Log, WindowTitle,
                                    "The log could not be saved:" & vbCrLf & vbCrLf & ex.Message)
        End Try

    End Sub

    ''' <summary>
    ''' Reads a .buslog back. The colours are not in the file - they come from the
    ''' marker column, so they are restored simply by loading it.
    ''' </summary>
    Public Sub LoadLog(filePath As String)

        If Rows Is Nothing OrElse Not File.Exists(filePath) Then Exit Sub

        Try
            Dim lines() As String = File.ReadAllLines(filePath)

            Rows.Clear()

            For index As Integer = 0 To lines.Length - 1

                ' First line is the header.
                If index = 0 Then Continue For
                If lines(index).Trim().Length = 0 Then Continue For

                Dim fields() As String = SplitCsv(lines(index))

                Rows.Add(New EventLogRow With {
                    .Marker = Field(fields, 1),
                    .Host = Field(fields, 2),
                    .Reg = Field(fields, 3),
                    .Value = Field(fields, 4),
                    .Comment = Field(fields, 5)
                })

            Next

            If Rows.Count = 0 Then Rows.Add(New EventLogRow)

            LogFilePath = filePath
            RefreshTitle()
            AppCore.SetStatus("Loaded " & Path.GetFileName(filePath))

        Catch ex As Exception
            MessagePrompt.ShowError(Log, WindowTitle,
                                    "The log could not be read:" & vbCrLf & vbCrLf & ex.Message)
        End Try

    End Sub

    Private Function Field(fields() As String, index As Integer) As String

        If index > fields.Length - 1 Then Return String.Empty

        Return fields(index)

    End Function

    ''' <summary>Wraps a field in quotes only when it needs them.</summary>
    Private Function Quote(value As String) As String

        Dim text As String = If(value, String.Empty)

        If text.IndexOfAny({","c, """"c, ControlChars.Cr, ControlChars.Lf}) < 0 Then Return text

        Return """" & text.Replace("""", """""") & """"

    End Function

    ''' <summary>Splits one CSV line, honouring quoted fields and doubled quotes.</summary>
    Private Function SplitCsv(line As String) As String()

        Dim fields As New List(Of String)
        Dim current As New StringBuilder()
        Dim inQuotes As Boolean = False
        Dim index As Integer = 0

        While index < line.Length

            Dim character As Char = line(index)

            If inQuotes Then

                If character = """"c Then
                    ' A doubled quote inside a quoted field is one literal quote.
                    If index < line.Length - 1 AndAlso line(index + 1) = """"c Then
                        current.Append(""""c)
                        index += 1
                    Else
                        inQuotes = False
                    End If
                Else
                    current.Append(character)
                End If

            ElseIf character = """"c Then
                inQuotes = True

            ElseIf character = ","c Then
                fields.Add(current.ToString())
                current.Clear()

            Else
                current.Append(character)
            End If

            index += 1

        End While

        fields.Add(current.ToString())

        Return fields.ToArray()

    End Function


    ' ========================================================================
    '  Replay
    ' ========================================================================

    Public Sub ReplayLine()

        ' TODO: put the current line on the bus.
        AppCore.ReportNotImplemented("Replay")

    End Sub

    Public Sub ReplayBlock()

        ' TODO: put the selected run of lines on the bus.
        AppCore.ReportNotImplemented("Replay Block")

    End Sub

    Public Sub ReplayAll()

        ' TODO: put the whole log on the bus.
        AppCore.ReportNotImplemented("Replay All")

    End Sub

End Module
