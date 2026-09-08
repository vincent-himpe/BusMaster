' ============================================================================
'  Modules\CommandCore.vb
'
'  The Command window: a typed way at the same registers the panels show.
'
'  Everything is addressed by name, in three levels:
'
'      device.register              a whole byte
'      device.register.bit          one bit, by number 7..0 or by its caption
'
'  The names are the ones the event log prints - the function name of the panel,
'  the register's name, the bit's caption - all with spaces written as
'  underscores. That is deliberate: a line read off the log can be typed straight
'  back in here.
'
'  Two operators:
'
'      = value      assign. A register takes 0-255 (or 0x..); a bit takes 1, 0
'                   or "/" to invert it.
'      ?            read it back and print what came back.
'
'  Nothing here touches a register directly. An assignment ends in RequestWrite
'  and a query in RequestRead, so a typed command reaches the hardware down the
'  same path as a mouse click and lands in the log the same way.
' ============================================================================

Imports System.Globalization
Imports System.Text
Imports System.Windows.Documents
Imports System.Windows.Threading

Public Module CommandCore

    Private Const WindowTitle As String = "Command"

    ''' <summary>What separates the levels of a name.</summary>
    Private Const LevelSeparator As Char = "."c

    Private Const AssignOperator As Char = "="c
    Private Const QueryOperator As Char = "?"c

    ''' <summary>Written after a command that could not be made sense of.</summary>
    Private Const SyntaxError As String = " <- Syntax Error"

    ''' <summary>A bit assignment of "/" means invert whatever is there.</summary>
    Private Const InvertValue As String = "/"

    ''' <summary>A line starting with this is a note to the reader and nothing else.</summary>
    Private Const CommentMark As Char = "'"c

    Private Window As CommandWindow

    ''' <summary>
    ''' The history line the caret is on. Held so it can be un-highlighted and
    ''' coloured in once the caret moves off it.
    ''' </summary>
    Private ActiveLine As Paragraph

    ''' <summary>Stops the colouring pass setting itself off again.</summary>
    Private Painting As Boolean = False


    ' ========================================================================
    '  Start-up and shutdown
    ' ========================================================================

    ''' <summary>
    ''' The parameter is deliberately not called "window": VB is case insensitive,
    ''' so it would be the very same identifier as the field above and "Window =
    ''' window" would quietly assign the parameter to itself.
    ''' </summary>
    Public Sub Attach(created As CommandWindow)

        Window = created

        WindowTheme.ApplyDarkTitleBar(Window)

        Window.rtb_CommandHistory.Document.Blocks.Clear()

        For Each line As String In {
            "' device.register = 22        assign a byte, or 0x16",
            "' device.register.3 = 1       one bit - 1, 0, or / to invert",
            "' device.register.Update = /  bits can be named as well as numbered",
            "' device.register ?           read it back",
            "' Anything after the value is a note and is ignored.",
            "' Enter runs the line the caret is on, Ctrl+Enter makes a new one."}

            AddHistory(line)

        Next

    End Sub

    ''' <summary>
    ''' Put away rather than thrown away, like the log: what has been typed and run
    ''' is still there next time it is opened.
    ''' </summary>
    Public Sub HandleWindowClosing(e As ComponentModel.CancelEventArgs)

        If Window Is Nothing Then Exit Sub

        e.Cancel = True
        Window.Hide()

    End Sub

    Public Sub ShowCommand()

        If Window Is Nothing Then Exit Sub

        Window.Show()

        If Window.WindowState = WindowState.Minimized Then Window.WindowState = WindowState.Normal

        Window.Activate()

        ' There is only one thing to do in this window, so put the caret there.
        Window.Dispatcher.BeginInvoke(DispatcherPriority.Input,
            New Action(Sub() Window.txt_ActiveCommand.Focus()))

    End Sub


    ' ========================================================================
    '  Typing
    ' ========================================================================

    ''' <summary>
    ''' Enter runs the line and leaves it alone - pressing it again runs the same
    ''' thing again, which is what makes this usable as a poke-and-watch tool.
    ''' Escape is what clears it.
    ''' </summary>
    Public Sub HandleKey(e As KeyEventArgs)

        If Window Is Nothing Then Exit Sub

        Select Case e.Key

            Case Key.Enter
                CloseCompletions()
                RunActiveCommand()
                e.Handled = True

            Case Key.Tab
                ' Tab never moves focus in this box; there is nowhere useful to go.
                If CompletionsShowing() Then AcceptCompletion()
                e.Handled = True

            Case Key.Escape
                ' The list first, if it is up - otherwise the line.
                If CompletionsShowing() Then
                    CloseCompletions()
                Else
                    Window.txt_ActiveCommand.Clear()
                End If
                e.Handled = True

            Case Key.Down
                If CompletionsShowing() Then
                    MoveCompletion(1)
                    e.Handled = True
                End If

            Case Key.Up
                If CompletionsShowing() Then
                    MoveCompletion(-1)
                    e.Handled = True
                End If

        End Select

    End Sub

    Private Sub RunActiveCommand()

        Dim typed As String = If(Window.txt_ActiveCommand.Text, String.Empty).Trim()
        If typed.Length = 0 Then Exit Sub

        ' A comment typed here is just filed away, which is a tidy way to annotate
        ' a run without leaving the keyboard.
        If IsComment(typed) Then
            AddHistory(typed)
            Exit Sub
        End If

        AddHistory(typed & Run(typed))

    End Sub

    ''' <summary>Puts a line on the end of the history, coloured, and scrolls to it.</summary>
    Private Sub AddHistory(line As String)

        Dim history As RichTextBox = Window.rtb_CommandHistory

        Painting = True

        Dim block As New Paragraph()
        Paint(block, line)
        history.Document.Blocks.Add(block)

        Painting = False

        history.ScrollToEnd()

    End Sub


    ' ========================================================================
    '  The history as a scratch pad
    '
    '  Editable, so it doubles as a place to keep and rearrange the commands
    '  worth running again. Running one from here deliberately writes nothing
    '  back: what the history holds is the user's to curate, and the effect of a
    '  read is visible on the panels anyway.
    ' ========================================================================

    Public Sub HandleHistoryKey(e As KeyEventArgs)

        If Window Is Nothing OrElse e.Key <> Key.Enter Then Exit Sub

        Dim control As Boolean = (Keyboard.Modifiers And ModifierKeys.Control) = ModifierKeys.Control

        If control Then
            ' The only way to get a new line, since plain Enter is spoken for.
            Dim history As RichTextBox = Window.rtb_CommandHistory
            history.CaretPosition = history.CaretPosition.InsertParagraphBreak()
        Else
            RunHistoryLine()
        End If

        e.Handled = True

    End Sub

    ''' <summary>
    ''' Runs the line the caret is on and steps to the next one, so holding Enter
    ''' walks a block of commands in order. Blank and comment lines are stepped over
    ''' rather than stopping the walk.
    ''' </summary>
    Private Sub RunHistoryLine()

        Dim block As Paragraph = CaretLine()
        If block Is Nothing Then Exit Sub

        Dim line As String = LineText(block).Trim()

        If line.Length > 0 AndAlso Not IsComment(line) Then

            ' Nothing goes back into the history. The outcome goes to the main
            ' window's status bar instead, which is the one place a result can be
            ' said without writing in the user's scratch pad.
            AppCore.SetStatus(Command(line) & Run(line))

        End If

        StepToNextLine(block)

    End Sub

    ''' <summary>
    ''' Moves the caret to the start of the following line. On the last line it
    ''' stays put, so a final Enter runs the same thing again rather than falling
    ''' off the end.
    ''' </summary>
    Private Sub StepToNextLine(block As Paragraph)

        Dim following As Paragraph = TryCast(block.NextBlock, Paragraph)
        If following Is Nothing Then Exit Sub

        Window.rtb_CommandHistory.CaretPosition = following.ContentStart

        ScrollCaretIntoView()

    End Sub

    ''' <summary>
    ''' Keeps the caret on screen while walking a list longer than the box.
    ''' </summary>
    Private Sub ScrollCaretIntoView()

        Dim history As RichTextBox = Window.rtb_CommandHistory
        Dim caret As TextPointer = history.CaretPosition
        If caret Is Nothing Then Exit Sub

        Dim spot As Rect = caret.GetCharacterRect(LogicalDirection.Forward)
        If spot.IsEmpty Then Exit Sub

        If spot.Bottom > history.ViewportHeight Then
            history.ScrollToVerticalOffset(history.VerticalOffset + spot.Bottom - history.ViewportHeight)
        ElseIf spot.Top < 0 Then
            history.ScrollToVerticalOffset(history.VerticalOffset + spot.Top)
        End If

    End Sub

    ''' <summary>
    ''' The part of a line that is the command: up to the value, or up to the "?".
    ''' What follows is the user's note or the answer from the last run, and showing
    ''' either of those next to a fresh answer only reads as two answers.
    ''' </summary>
    Private Function Command(line As String) As String

        Dim assignAt As Integer = line.IndexOf(AssignOperator)
        Dim queryAt As Integer = line.IndexOf(QueryOperator)

        If assignAt >= 0 AndAlso (queryAt < 0 OrElse assignAt < queryAt) Then

            Dim tail As String = line.Substring(assignAt + 1)
            Dim lead As Integer = tail.Length - tail.TrimStart().Length

            Return line.Substring(0, assignAt + 1 + lead + FirstWord(tail).Length)

        End If

        If queryAt >= 0 Then Return line.Substring(0, queryAt + 1)

        Return line

    End Function

    ''' <summary>
    ''' Follows the caret: the line it is on gets the blue backdrop, and the line it
    ''' has just left gets painted in.
    ''' </summary>
    Public Sub HistoryCaretMoved()

        If Window Is Nothing OrElse Painting Then Exit Sub

        Dim block As Paragraph = CaretLine()
        If block Is ActiveLine Then Exit Sub

        Painting = True

        If ActiveLine IsNot Nothing Then
            ActiveLine.Background = Nothing

            ' Colouring while the caret sits on a line would fight whatever is being
            ' typed, so it waits until the caret has moved off.
            Paint(ActiveLine, LineText(ActiveLine))
        End If

        ActiveLine = block

        If ActiveLine IsNot Nothing Then ActiveLine.Background = Ink("Brush_Command_Active")

        Painting = False

    End Sub

    Private Function CaretLine() As Paragraph

        Dim caret As TextPointer = Window.rtb_CommandHistory.CaretPosition
        If caret Is Nothing Then Return Nothing

        Return caret.Paragraph

    End Function

    Private Function LineText(block As Paragraph) As String

        Return New TextRange(block.ContentStart, block.ContentEnd).Text

    End Function

    Private Function IsComment(line As String) As Boolean

        Return line.TrimStart().StartsWith(CommentMark)

    End Function


    ' ========================================================================
    '  Running a command
    '
    '  Returns what should follow the command in the history - the answer to a
    '  query, the syntax error, or nothing at all for an assignment that worked.
    ' ========================================================================

    ''' <summary>
    ''' Reading stops as soon as the command is complete, and whatever follows is
    ''' left alone. That is what lets a line be run again with its own answer still
    ''' written after it, and what lets the user write a note on the end.
    ''' </summary>
    Private Function Run(command As String) As String

        Dim assignAt As Integer = command.IndexOf(AssignOperator)
        Dim queryAt As Integer = command.IndexOf(QueryOperator)

        ' Whichever comes first decides what kind of line this is - a "?" inside the
        ' answer to an earlier read must not turn an assignment into a query.
        If assignAt >= 0 AndAlso (queryAt < 0 OrElse assignAt < queryAt) Then
            Return Assign(command.Substring(0, assignAt), FirstWord(command.Substring(assignAt + 1)))
        End If

        If queryAt >= 0 Then
            Return Query(command.Substring(0, queryAt))
        End If

        ' A name on its own says nothing about what to do with it.
        Return SyntaxError

    End Function

    ''' <summary>The first run of non-blank characters, or an empty string.</summary>
    Private Function FirstWord(text As String) As String

        Dim trimmed As String = text.TrimStart()
        If trimmed.Length = 0 Then Return String.Empty

        Dim ends As Integer = trimmed.IndexOf(" "c)

        Return If(ends < 0, trimmed, trimmed.Substring(0, ends))

    End Function

    Private Function Assign(path As String, value As String) As String

        Dim target As CommandTarget = Resolve(path)
        If target Is Nothing Then Return SyntaxError

        Dim wanted As String = value.Trim()

        If target.IsWholeRegister Then

            Dim parsed As Integer = ParseByte(wanted)
            If parsed < 0 Then Return SyntaxError

            target.Editor.Value = CByte(parsed)

        Else

            Dim mask As Integer = 1 << target.BitNumber
            Dim current As Integer = CInt(target.Editor.Value)

            Select Case wanted
                Case "1"
                    current = current Or mask
                Case "0"
                    current = current And (Not mask)
                Case InvertValue
                    current = current Xor mask
                Case Else
                    Return SyntaxError
            End Select

            ' Read, mask, write back - the same arithmetic a bit click does, on the
            ' byte the panel is already holding.
            target.Editor.Value = CByte(current And &HFF)

        End If

        ' Setting Value in code raises nothing, so the write has to be asked for.
        target.Editor.RequestWrite()

        Return String.Empty

    End Function

    Private Function Query(path As String) As String

        Dim target As CommandTarget = Resolve(path.Trim())
        If target Is Nothing Then Return SyntaxError

        ' A query really reads. What comes back is what gets printed.
        target.Editor.RequestRead()

        Dim value As Integer = CInt(target.Editor.Value)

        If target.IsWholeRegister Then
            Return " 0x" & value.ToString("X2", CultureInfo.InvariantCulture) &
                   " , " & value.ToString(CultureInfo.InvariantCulture)
        End If

        Dim isSet As Boolean = ((value >> target.BitNumber) And 1) = 1

        Return If(isSet, " TRUE", " FALSE")

    End Function

    ''' <summary>
    ''' A byte written as decimal or as 0x hex, or -1 if it is neither or will not
    ''' fit in one.
    ''' </summary>
    Private Function ParseByte(text As String) As Integer

        If text.Length = 0 Then Return -1

        Dim parsed As Integer
        Dim ok As Boolean

        If text.StartsWith("0x", StringComparison.OrdinalIgnoreCase) Then
            ok = Integer.TryParse(text.Substring(2), NumberStyles.HexNumber,
                                  CultureInfo.InvariantCulture, parsed)
        Else
            ok = Integer.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, parsed)
        End If

        If Not ok OrElse parsed < 0 OrElse parsed > 255 Then Return -1

        Return parsed

    End Function


    ' ========================================================================
    '  Colouring
    '
    '  Each of the three levels of a name gets its own pastel, so the shape of a
    '  line is readable before any of it is. Everything structural - the dots,
    '  the operator - stays white.
    ' ========================================================================

    ''' <summary>Rebuilds a line's runs from its text.</summary>
    Private Sub Paint(block As Paragraph, line As String)

        block.Inlines.Clear()

        For Each piece As CommandPiece In SplitLine(line)
            block.Inlines.Add(New Run(piece.Text) With {.Foreground = piece.Ink})
        Next

    End Sub

    ''' <summary>Breaks a line into the stretches that get their own colour.</summary>
    Private Function SplitLine(line As String) As List(Of CommandPiece)

        Dim pieces As New List(Of CommandPiece)
        If line Is Nothing OrElse line.Length = 0 Then Return pieces

        ' The complaint on the end of a line is an answer like any other, so it is
        ' set aside here and painted red rather than being taken for a note.
        If line.EndsWith(SyntaxError, StringComparison.Ordinal) Then

            Dim withoutComplaint As String = line.Substring(0, line.Length - SyntaxError.Length)

            pieces.AddRange(SplitLine(withoutComplaint))
            pieces.Add(New CommandPiece(SyntaxError, Ink("Brush_Command_Result")))

            Return pieces

        End If

        ' Leading blanks belong to nothing in particular.
        Dim indent As Integer = line.Length - line.TrimStart().Length
        If indent > 0 Then pieces.Add(New CommandPiece(line.Substring(0, indent), Ink("Brush_Command_Plain")))

        Dim rest As String = line.Substring(indent)
        If rest.Length = 0 Then Return pieces

        If rest(0) = CommentMark Then
            pieces.Add(New CommandPiece(rest, Ink("Brush_Command_Comment")))
            Return pieces
        End If

        Dim opAt As Integer = rest.IndexOfAny({AssignOperator, QueryOperator})

        AddName(pieces, If(opAt < 0, rest, rest.Substring(0, opAt)))

        If opAt < 0 Then Return pieces

        Dim operatorUsed As Char = rest(opAt)
        pieces.Add(New CommandPiece(operatorUsed.ToString(), Ink("Brush_Command_Plain")))

        Dim tail As String = rest.Substring(opAt + 1)
        If tail.Length = 0 Then Return pieces

        ' After a "?" whatever is there is the answer from the last time it ran.
        If operatorUsed = QueryOperator Then
            pieces.Add(New CommandPiece(tail, Ink("Brush_Command_Result")))
            Return pieces
        End If

        ' After an "=", the first word is the value and the rest is the user's note.
        Dim lead As Integer = tail.Length - tail.TrimStart().Length
        If lead > 0 Then pieces.Add(New CommandPiece(tail.Substring(0, lead), Ink("Brush_Command_Plain")))

        Dim body As String = tail.Substring(lead)
        If body.Length = 0 Then Return pieces

        Dim ends As Integer = body.IndexOf(" "c)

        If ends < 0 Then
            pieces.Add(New CommandPiece(body, Ink("Brush_Command_Number")))
        Else
            pieces.Add(New CommandPiece(body.Substring(0, ends), Ink("Brush_Command_Number")))
            pieces.Add(New CommandPiece(body.Substring(ends), Ink("Brush_Command_Comment")))
        End If

        Return pieces

    End Function

    ''' <summary>Device, register, bit - each in its own colour, dots left white.</summary>
    Private Sub AddName(pieces As List(Of CommandPiece), name As String)

        Dim levels() As String = {"Brush_Command_Device", "Brush_Command_Register", "Brush_Command_Bit"}
        Dim segments() As String = name.Split(LevelSeparator)

        For index As Integer = 0 To segments.Length - 1

            If index > 0 Then pieces.Add(New CommandPiece(".", Ink("Brush_Command_Plain")))

            Dim key As String = If(index < levels.Length, levels(index), "Brush_Command_Plain")

            If segments(index).Length > 0 Then pieces.Add(New CommandPiece(segments(index), Ink(key)))

        Next

    End Sub

    ''' <summary>A theme brush by key, or white if the theme is not there.</summary>
    Private Function Ink(key As String) As Brush

        If System.Windows.Application.Current IsNot Nothing Then
            Dim found As Object = System.Windows.Application.Current.TryFindResource(key)
            If TypeOf found Is Brush Then Return DirectCast(found, Brush)
        End If

        Return Brushes.White

    End Function


    ' ========================================================================
    '  Resolving a name
    ' ========================================================================

    ''' <summary>
    ''' Turns "device.register" or "device.register.bit" into the control it names,
    ''' or Nothing if no such thing is on screen. Two panels sharing a function name
    ''' is the user's business; the first one wins.
    ''' </summary>
    Private Function Resolve(path As String) As CommandTarget

        Dim parts() As String = path.Trim().Split(LevelSeparator)

        If parts.Length < 2 OrElse parts.Length > 3 Then Return Nothing

        Dim panel As DevicePanel = FindPanel(parts(0))
        If panel Is Nothing Then Return Nothing

        Dim editor As BitFieldEditor = FindRegister(panel, parts(1))
        If editor Is Nothing Then Return Nothing

        If parts.Length = 2 Then Return New CommandTarget(editor, -1)

        Dim bit As Integer = FindBit(editor, parts(2))
        If bit < 0 Then Return Nothing

        Return New CommandTarget(editor, bit)

    End Function

    Private Function FindPanel(name As String) As DevicePanel

        Dim wanted As String = TextTools.Symbolic(name)
        If wanted.Length = 0 Then Return Nothing

        For Each panel As DevicePanel In AppCore.DevicePanels()
            If Same(TextTools.Symbolic(panel.FunctionName), wanted) Then Return panel
        Next

        Return Nothing

    End Function

    Private Function FindRegister(panel As DevicePanel, name As String) As BitFieldEditor

        Dim wanted As String = TextTools.Symbolic(name)
        If wanted.Length = 0 Then Return Nothing

        For Each editor As BitFieldEditor In panel.Registers()
            If Same(TextTools.Symbolic(editor.Registername), wanted) Then Return editor
        Next

        Return Nothing

    End Function

    ''' <summary>A bit by its number, 7 down to 0, or by the caption it carries.</summary>
    Private Function FindBit(editor As BitFieldEditor, name As String) As Integer

        Dim wanted As String = name.Trim()
        If wanted.Length = 0 Then Return -1

        Dim number As Integer

        If Integer.TryParse(wanted, NumberStyles.None, CultureInfo.InvariantCulture, number) Then
            If number < 0 OrElse number > BitFieldEditor.BitCount - 1 Then Return -1
            Return number
        End If

        Return editor.BitNumberFor(wanted)

    End Function

    Private Function Same(left As String, right As String) As Boolean

        Return String.Equals(left, right, StringComparison.OrdinalIgnoreCase)

    End Function


    ' ========================================================================
    '  Completion
    '
    '  Offers whatever fits the level the caret is in. Driven from the keyboard
    '  only - the list never takes focus, so the caret stays where it is.
    ' ========================================================================

    Public Sub RefreshCompletions()

        If Window Is Nothing Then Exit Sub

        Dim box As TextBox = Window.txt_ActiveCommand
        Dim ahead As String = box.Text.Substring(0, Math.Min(box.CaretIndex, box.Text.Length))

        ' Past an operator there are no more names to offer.
        If ahead.IndexOf(AssignOperator) >= 0 OrElse ahead.IndexOf(QueryOperator) >= 0 Then
            CloseCompletions()
            Exit Sub
        End If

        Dim parts() As String = ahead.Split(LevelSeparator)
        Dim candidates As List(Of String) = CandidatesFor(parts)

        If candidates.Count = 0 Then
            CloseCompletions()
            Exit Sub
        End If

        Window.lst_Completions.ItemsSource = candidates
        Window.lst_Completions.SelectedIndex = 0
        Window.pop_Completions.IsOpen = True

    End Sub

    ''' <summary>
    ''' What could follow at the level being typed, narrowed to what has been typed
    ''' of it so far. Level one is devices, two is that device's registers, three is
    ''' that register's named bits - and there is no level four.
    ''' </summary>
    Private Function CandidatesFor(parts() As String) As List(Of String)

        Dim found As New List(Of String)

        If parts.Length < 1 OrElse parts.Length > 3 Then Return found

        Dim prefix As String = TextTools.Symbolic(parts(parts.Length - 1))

        Select Case parts.Length

            Case 1
                For Each panel As DevicePanel In AppCore.DevicePanels()
                    Offer(found, TextTools.Symbolic(panel.FunctionName), prefix)
                Next

            Case 2
                Dim panel As DevicePanel = FindPanel(parts(0))
                If panel Is Nothing Then Return found

                For Each editor As BitFieldEditor In panel.Registers()
                    Offer(found, TextTools.Symbolic(editor.Registername), prefix)
                Next

            Case 3
                Dim panel As DevicePanel = FindPanel(parts(0))
                If panel Is Nothing Then Return found

                Dim editor As BitFieldEditor = FindRegister(panel, parts(1))
                If editor Is Nothing Then Return found

                For Each caption As String In editor.BitNames()
                    Offer(found, TextTools.Symbolic(caption), prefix)
                Next

        End Select

        Return found

    End Function

    ''' <summary>Adds a candidate if it fits what has been typed and is not already there.</summary>
    Private Sub Offer(found As List(Of String), candidate As String, prefix As String)

        If candidate.Length = 0 Then Exit Sub
        If Not candidate.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) Then Exit Sub
        If found.Contains(candidate) Then Exit Sub

        found.Add(candidate)

    End Sub

    Public Function CompletionsShowing() As Boolean

        Return Window IsNot Nothing AndAlso Window.pop_Completions.IsOpen

    End Function

    Public Sub CloseCompletions()

        If Window Is Nothing Then Exit Sub

        Window.pop_Completions.IsOpen = False

    End Sub

    Private Sub MoveCompletion(offset As Integer)

        Dim list As ListBox = Window.lst_Completions
        If list.Items.Count = 0 Then Exit Sub

        Dim target As Integer = list.SelectedIndex + offset

        If target < 0 Then target = list.Items.Count - 1
        If target > list.Items.Count - 1 Then target = 0

        list.SelectedIndex = target
        list.ScrollIntoView(list.SelectedItem)

    End Sub

    ''' <summary>
    ''' Puts the highlighted name in place of the level being typed, leaving the
    ''' caret at the end of it so the next "." can be typed.
    ''' </summary>
    Private Sub AcceptCompletion()

        Dim chosen As String = TryCast(Window.lst_Completions.SelectedItem, String)
        If chosen Is Nothing Then Exit Sub

        Dim box As TextBox = Window.txt_ActiveCommand
        Dim caret As Integer = Math.Min(box.CaretIndex, box.Text.Length)

        ' The level being typed starts after the last separator before the caret.
        Dim levelStart As Integer = box.Text.LastIndexOf(LevelSeparator, Math.Max(caret - 1, 0)) + 1
        If caret = 0 Then levelStart = 0

        Dim rebuilt As New StringBuilder()
        rebuilt.Append(box.Text.Substring(0, levelStart))
        rebuilt.Append(chosen)
        rebuilt.Append(box.Text.Substring(caret))

        box.Text = rebuilt.ToString()
        box.CaretIndex = levelStart + chosen.Length

        CloseCompletions()

    End Sub

End Module


''' <summary>One stretch of a history line and the colour it is written in.</summary>
Friend Class CommandPiece

    Public ReadOnly Text As String
    Public ReadOnly Ink As Brush

    Public Sub New(text As String, ink As Brush)

        Me.Text = text
        Me.Ink = ink

    End Sub

End Class


''' <summary>
''' What a name resolved to: a register, and which bit of it - or the whole byte.
''' </summary>
Friend Class CommandTarget

    Public ReadOnly Editor As BitFieldEditor

    ''' <summary>7 down to 0, or -1 for the register as a whole.</summary>
    Public ReadOnly BitNumber As Integer

    Public Sub New(editor As BitFieldEditor, bitNumber As Integer)

        Me.Editor = editor
        Me.BitNumber = bitNumber

    End Sub

    Public ReadOnly Property IsWholeRegister As Boolean
        Get
            Return BitNumber < 0
        End Get
    End Property

End Class
