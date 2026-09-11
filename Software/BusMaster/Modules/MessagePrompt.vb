' ============================================================================
'  Modules\MessagePrompt.vb
'
'  "Here is something you should know", and "are you sure?" - in this program's
'  colours.
'
'  The companion to StringPrompt: that one asks for a string, this one says a
'  piece or puts a question. Both exist for the same reason, which is that the
'  system's own boxes draw themselves in light grey and stand out against
'  everything else here.
'
'  Four faces, and the glyph is the only thing on the form that carries a
'  colour:
'
'      Info      blue circle with an i       something worth knowing
'      Warning   yellow triangle             something worth thinking twice about
'      Error     red circle with a cross     something that did not happen
'      Question  green circle with a ?       something only the user can decide
'
'  Pick by what the message is, not by how loud it should be. A refusal the user
'  could not have known about is information; one that undoes something they
'  meant is a warning; one that could not be carried out is an error.
'
'  Four button sets, and the answer comes back as a MessageBoxResult so that a
'  call site reads the same after being moved over from MessageBox.Show:
'
'      ShowInfo / ShowWarning / ShowError   OK
'      AskOkCancel                          OK, Cancel
'      AskYesNo                             Yes, No
'      AskYesNoCancel                       Yes, No, Cancel
'
'  Escape always answers the way the safe button does - Cancel where there is
'  one, No where there is not, and OK when OK is all there is.
' ============================================================================

Imports System.Windows.Input

''' <summary>Which buttons a message box puts up.</summary>
Public Enum MessageButtons
    buttons_Ok = 0
    buttons_OkCancel = 1
    buttons_YesNo = 2
    buttons_YesNoCancel = 3
End Enum


Public Module MessagePrompt

    ' Written as codepoints rather than as literal characters so the source stays
    ' readable in an editor with no icon font.
    Private ReadOnly InfoGlyph As String = ChrW(&HE946)
    Private ReadOnly WarningGlyph As String = ChrW(&HE7BA)
    Private ReadOnly ErrorGlyph As String = ChrW(&HEA39)
    Private ReadOnly QuestionGlyph As String = ChrW(&HE9CE)


    ' ========================================================================
    '  Saying a piece
    ' ========================================================================

    ''' <summary>Something worth knowing. Blue.</summary>
    Public Sub ShowInfo(owner As Window, title As String, message As String)

        Say(owner, title, message, InfoGlyph, "Brush_Message_Info", MessageButtons.buttons_Ok)

    End Sub

    ''' <summary>Something worth thinking twice about. Yellow.</summary>
    Public Sub ShowWarning(owner As Window, title As String, message As String)

        Say(owner, title, message, WarningGlyph, "Brush_Message_Warning", MessageButtons.buttons_Ok)

    End Sub

    ''' <summary>Something that did not happen. Red.</summary>
    Public Sub ShowError(owner As Window, title As String, message As String)

        Say(owner, title, message, ErrorGlyph, "Brush_Message_Error", MessageButtons.buttons_Ok)

    End Sub


    ' ========================================================================
    '  Putting a question
    '
    '  All three wear the green question mark: what is being asked differs, but
    '  the fact that the user is being asked does not.
    ' ========================================================================

    ''' <summary>Go ahead, or leave it? Returns OK or Cancel.</summary>
    Public Function AskOkCancel(owner As Window, title As String, message As String) As MessageBoxResult

        Return Say(owner, title, message, QuestionGlyph, "Brush_Message_Question",
                   MessageButtons.buttons_OkCancel)

    End Function

    ''' <summary>One thing or the other, with no way out. Returns Yes or No.</summary>
    Public Function AskYesNo(owner As Window, title As String, message As String) As MessageBoxResult

        Return Say(owner, title, message, QuestionGlyph, "Brush_Message_Question",
                   MessageButtons.buttons_YesNo)

    End Function

    ''' <summary>
    ''' One thing, the other, or neither - the shape of "save first?". Returns Yes,
    ''' No or Cancel.
    ''' </summary>
    Public Function AskYesNoCancel(owner As Window, title As String, message As String) As MessageBoxResult

        Return Say(owner, title, message, QuestionGlyph, "Brush_Message_Question",
                   MessageButtons.buttons_YesNoCancel)

    End Function


    ' ========================================================================
    '  The box itself
    ' ========================================================================

    ''' <summary>
    ''' Puts the message up and waits for an answer. Modal, like the box it
    ''' replaces.
    ''' </summary>
    Private Function Say(owner As Window, title As String, message As String,
                         glyph As String, brushKey As String,
                         buttons As MessageButtons) As MessageBoxResult

        Dim dialog As New MessagePromptWindow

        If owner IsNot Nothing Then dialog.Owner = owner

        ' The box sizes itself to what it has been given, and the F1 help is long
        ' enough to reach past the bottom of a laptop screen. Capped here rather
        ' than in the XAML because it depends on the screen the program is on, and
        ' the message then scrolls inside the box instead of falling off it.
        dialog.MaxHeight = Math.Max(SystemParameters.WorkArea.Height - 80, 240)

        dialog.Title = If(title, String.Empty)
        dialog.lbl_Message.Text = If(message, String.Empty)
        dialog.lbl_Icon.Text = glyph

        Dim ink As Brush = TryCast(dialog.TryFindResource(brushKey), Brush)
        If ink IsNot Nothing Then dialog.lbl_Icon.Foreground = ink

        Arrange(dialog, buttons)

        ModalShade.Cover(dialog)

        Try
            dialog.ShowDialog()
        Finally
            ModalShade.Uncover()
        End Try

        Return dialog.Answer

    End Function

    ''' <summary>
    ''' Shows the buttons this mode calls for and hides the rest, decides which one
    ''' Enter presses, and settles what Escape means.
    '''
    ''' Collapsed rather than hidden: a button that is not offered should take up no
    ''' room, so the row closes up round the ones that are.
    ''' </summary>
    Private Sub Arrange(dialog As MessagePromptWindow, buttons As MessageButtons)

        Dim wantsOk As Boolean = buttons = MessageButtons.buttons_Ok OrElse
                                 buttons = MessageButtons.buttons_OkCancel

        Dim wantsYesNo As Boolean = buttons = MessageButtons.buttons_YesNo OrElse
                                    buttons = MessageButtons.buttons_YesNoCancel

        Dim wantsCancel As Boolean = buttons = MessageButtons.buttons_OkCancel OrElse
                                     buttons = MessageButtons.buttons_YesNoCancel

        dialog.btn_Ok.Visibility = If(wantsOk, Visibility.Visible, Visibility.Collapsed)
        dialog.btn_Yes.Visibility = If(wantsYesNo, Visibility.Visible, Visibility.Collapsed)
        dialog.btn_No.Visibility = If(wantsYesNo, Visibility.Visible, Visibility.Collapsed)
        dialog.btn_Cancel.Visibility = If(wantsCancel, Visibility.Visible, Visibility.Collapsed)

        ' Enter presses the one that gets on with it.
        dialog.btn_Ok.IsDefault = wantsOk
        dialog.btn_Yes.IsDefault = wantsYesNo

        ' Escape answers the safest way out there is.
        If wantsCancel Then
            dialog.EscapeAnswer = MessageBoxResult.Cancel
        ElseIf wantsYesNo Then
            dialog.EscapeAnswer = MessageBoxResult.No
        Else
            dialog.EscapeAnswer = MessageBoxResult.OK
        End If

    End Sub

    Public Sub Attach(dialog As MessagePromptWindow)

        WindowTheme.ApplyDarkTitleBar(dialog)

        ' The button Enter would press takes the focus, so the space bar agrees with
        ' Enter and neither has to be aimed at anything first.
        If dialog.btn_Yes.IsDefault Then
            dialog.btn_Yes.Focus()
        Else
            dialog.btn_Ok.Focus()
        End If

    End Sub

    ''' <summary>Escape means whatever the safe button on this box means.</summary>
    Public Sub HandleKey(dialog As MessagePromptWindow, e As KeyEventArgs)

        If e.Key <> Key.Escape Then Exit Sub

        Answered(dialog, dialog.EscapeAnswer)
        e.Handled = True

    End Sub

    ''' <summary>The one way out: record the answer, then close.</summary>
    Public Sub Answered(dialog As MessagePromptWindow, answer As MessageBoxResult)

        dialog.Answer = answer
        dialog.DialogResult = True

    End Sub

End Module
