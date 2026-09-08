' ============================================================================
'  Modules\TerminalCore.vb
'
'  The Terminal window: a plain view of the traffic to and from the probe.
'
'  Two boxes, like the Command window. The transcript on top is read-only - only
'  what is sent and what comes back write there - and the line being typed sits
'  underneath. Colour says which way a line went and nothing else: light green
'  out, light blue back.
'
'  Sending goes through ProbeControl.ProbeSend and nowhere else; arriving lines
'  come here from ProbeControl.ProbeReceive. Neither knows anything about the
'  other's end of the wire.
' ============================================================================

Imports System.Globalization
Imports System.Windows.Documents
Imports System.Windows.Threading

Public Module TerminalCore

    ''' <summary>How many lines the transcript keeps, and how far that can be taken.</summary>
    Private Const DefaultBufferSize As Integer = 20
    Private Const SmallestBufferSize As Integer = 5
    Private Const LargestBufferSize As Integer = 200

    Private Window As TerminalWindow

    Private BufferSize As Integer = DefaultBufferSize


    ' ========================================================================
    '  Start-up and shutdown
    ' ========================================================================

    ''' <summary>
    ''' The parameter is not called "window" on purpose - VB would read that as the
    ''' very same identifier as the field above.
    ''' </summary>
    Public Sub Attach(created As TerminalWindow)

        Window = created

        WindowTheme.ApplyDarkTitleBar(Window)

        ShowBufferSize()
        ShowProbeState()

    End Sub

    ''' <summary>
    ''' Says which probe the transcript is talking to. Keyed on the link actually
    ''' being up, not on what has been picked from the list - a port chosen but not
    ''' opened is still no probe as far as this window is concerned.
    ''' </summary>
    Public Sub ShowProbeState()

        If Window Is Nothing Then Exit Sub

        If ProbeControl.IsConnected Then
            Window.lbl_ProbePort.Text = ProbeControl.PortInUse
            Window.lbl_ProbePort.Foreground = Ink("Brush_Terminal_Received")
        Else
            Window.lbl_ProbePort.Text = "No Probe"
            Window.lbl_ProbePort.Foreground = Ink("Brush_Terminal_Notice")
        End If

    End Sub

    ''' <summary>Put away rather than thrown away, so the transcript survives.</summary>
    Public Sub HandleWindowClosing(e As ComponentModel.CancelEventArgs)

        If Window Is Nothing Then Exit Sub

        e.Cancel = True
        Window.Hide()

    End Sub

    Public Sub ShowTerminal()

        If Window Is Nothing Then Exit Sub

        Window.Show()

        If Window.WindowState = WindowState.Minimized Then Window.WindowState = WindowState.Normal

        Window.Activate()

        Window.Dispatcher.BeginInvoke(DispatcherPriority.Input,
            New Action(Sub() Window.txt_TerminalEntry.Focus()))

    End Sub


    ' ========================================================================
    '  Sending and receiving
    ' ========================================================================

    ''' <summary>
    ''' A line that has been sent but is still on the entry line. What happens to it
    ''' depends on what the user does next, so it is not cleared until then.
    ''' </summary>
    Private EntrySpent As Boolean = False

    Public Sub HandleEntryKey(e As KeyEventArgs)

        If Window Is Nothing Then Exit Sub

        Select Case e.Key

            Case Key.Enter
                SendEntry()
                e.Handled = True

            Case Key.Back, Key.Delete, Key.Left, Key.Right, Key.Home, Key.End
                ' Going back into the line rather than replacing it: "write 1",
                ' Enter, Backspace leaves "write " ready for a different argument.
                ' None of these can put a character in, so letting them through
                ' cannot damage a line the user meant to overwrite.
                EntrySpent = False

        End Select

    End Sub

    ''' <summary>
    ''' The first character typed after a line has gone replaces it. Handled here
    ''' rather than in the key handler because this is the one place that knows a
    ''' keystroke is really about to put text in the box.
    ''' </summary>
    Public Sub HandleEntryText(e As TextCompositionEventArgs)

        If Window Is Nothing OrElse Not EntrySpent Then Exit Sub

        EntrySpent = False
        Window.txt_TerminalEntry.Clear()

    End Sub

    ''' <summary>
    ''' Sends what has been typed and shows it in the transcript.
    '''
    ''' The entry line is deliberately left as it is. Clearing it here would throw
    ''' away a command the user very often wants to send again with one character
    ''' changed; instead it is marked as spent, and the next keystroke decides -
    ''' type, and it is replaced; backspace, and it is picked back up.
    ''' </summary>
    Private Sub SendEntry()

        Dim entry As TextBox = Window.txt_TerminalEntry

        Dim typed As String = If(entry.Text, String.Empty)
        If typed.Length = 0 Then Exit Sub

        ' Shown before it is sent, not after: a reply can arrive the moment the line
        ' goes out, and it has to land underneath the line that asked for it.
        AddLine(typed, "Brush_Terminal_Sent")

        ProbeControl.ProbeSend(typed)

        ' At the end, so a backspace takes off the last character rather than
        ' whatever the caret happened to be sitting in front of.
        entry.CaretIndex = entry.Text.Length
        EntrySpent = True

    End Sub

    ''' <summary>
    ''' A line has come back from the probe. Called by ProbeControl.ProbeReceive
    ''' once it has made whatever sense of it there is to make.
    ''' </summary>
    Public Sub ShowReceived(text As String)

        If Window Is Nothing Then Exit Sub

        AddLine(If(text, String.Empty), "Brush_Terminal_Received")

    End Sub

    ''' <summary>
    ''' Something the link has to say about itself rather than traffic it carried -
    ''' a port opened, a port closed. Grey: it happened, that is all.
    ''' </summary>
    Public Sub ShowNotice(text As String)

        If Window Is Nothing Then Exit Sub

        AddLine(If(text, String.Empty), "Brush_Terminal_Notice")

    End Sub

    ''' <summary>
    ''' Something that went wrong - a port that would not open, a write that failed.
    ''' Red, and only ever for this.
    ''' </summary>
    Public Sub ShowProblem(text As String)

        If Window Is Nothing Then Exit Sub

        AddLine(If(text, String.Empty), "Brush_Terminal_Problem")

    End Sub


    ' ========================================================================
    '  The transcript
    ' ========================================================================

    ''' <summary>Adds a line at the bottom and trims the top if that is now too many.</summary>
    Private Sub AddLine(text As String, inkKey As String)

        Dim view As RichTextBox = Window.rtb_TerminalHistory

        view.Document.Blocks.Add(New Paragraph(New Run(text) With {.Foreground = Ink(inkKey)}))

        TrimToBuffer()

        view.ScrollToEnd()

    End Sub

    ''' <summary>
    ''' Drops lines off the top until the transcript is within its buffer. Oldest
    ''' first, which is the only order that leaves what is on screen making sense.
    ''' </summary>
    Private Sub TrimToBuffer()

        If Window Is Nothing Then Exit Sub

        Dim lines As BlockCollection = Window.rtb_TerminalHistory.Document.Blocks

        While lines.Count > BufferSize
            lines.Remove(lines.FirstBlock)
        End While

    End Sub

    Public Sub ClearTranscript()

        If Window Is Nothing Then Exit Sub

        Window.rtb_TerminalHistory.Document.Blocks.Clear()

    End Sub

    ''' <summary>
    ''' Double-clicking a line puts it back on the entry line, ready to be sent
    ''' again or edited first. The transcript itself is left alone.
    ''' </summary>
    Public Sub CopyLineToEntry()

        If Window Is Nothing Then Exit Sub

        Dim caret As TextPointer = Window.rtb_TerminalHistory.Selection.Start
        If caret Is Nothing Then Exit Sub

        Dim line As Paragraph = caret.Paragraph
        If line Is Nothing Then Exit Sub

        Dim entry As TextBox = Window.txt_TerminalEntry

        entry.Text = New TextRange(line.ContentStart, line.ContentEnd).Text
        entry.CaretIndex = entry.Text.Length
        entry.Focus()

        ' Put there on purpose, so the next thing typed does not wipe it.
        EntrySpent = False

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
    '  Buffer size
    '
    '  Typed values are taken when the box is left or Enter is pressed, not on
    '  every keystroke - clamping half a number as it is typed would turn "150"
    '  into 5 at the first digit.
    ' ========================================================================

    Public Sub StepBufferSize(offset As Integer)

        ApplyBufferSize(BufferSize + offset)

    End Sub

    Public Sub CommitBufferSize()

        If Window Is Nothing Then Exit Sub

        Dim typed As Integer

        If Integer.TryParse(Window.txt_BufferSize.Text, NumberStyles.None,
                            CultureInfo.InvariantCulture, typed) Then
            ApplyBufferSize(typed)
        Else
            ' Nonsense in the box just goes back to what is actually in force.
            ShowBufferSize()
        End If

    End Sub

    Public Sub HandleBufferSizeKey(e As KeyEventArgs)

        Select Case e.Key

            Case Key.Enter
                CommitBufferSize()
                e.Handled = True

            Case Key.Up
                StepBufferSize(1)
                e.Handled = True

            Case Key.Down
                StepBufferSize(-1)
                e.Handled = True

        End Select

    End Sub

    ''' <summary>Digits only. Space never reaches PreviewTextInput, and cannot be typed anyway.</summary>
    Public Sub GuardBufferSizeEntry(e As TextCompositionEventArgs)

        For Each character As Char In If(e.Text, String.Empty)
            If character < "0"c OrElse character > "9"c Then
                e.Handled = True
                Exit Sub
            End If
        Next

    End Sub

    Private Sub ApplyBufferSize(wanted As Integer)

        BufferSize = Math.Max(SmallestBufferSize, Math.Min(LargestBufferSize, wanted))

        ShowBufferSize()
        TrimToBuffer()

    End Sub

    Private Sub ShowBufferSize()

        If Window Is Nothing Then Exit Sub

        Window.txt_BufferSize.Text = BufferSize.ToString(CultureInfo.InvariantCulture)

    End Sub

End Module
