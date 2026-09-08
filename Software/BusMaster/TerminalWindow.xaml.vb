' ============================================================================
'  TerminalWindow.xaml.vb
'
'  Event handlers ONLY - see Modules\TerminalCore.vb for everything they call.
' ============================================================================

Imports System.ComponentModel

Class TerminalWindow

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        TerminalCore.Attach(Me)
    End Sub

    Private Sub Window_Closing(sender As Object, e As CancelEventArgs)
        TerminalCore.HandleWindowClosing(e)
    End Sub

    Private Sub txt_TerminalEntry_PreviewKeyDown(sender As Object, e As KeyEventArgs)
        TerminalCore.HandleEntryKey(e)
    End Sub

    Private Sub txt_TerminalEntry_PreviewTextInput(sender As Object, e As TextCompositionEventArgs)
        TerminalCore.HandleEntryText(e)
    End Sub

    Private Sub rtb_TerminalHistory_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs)
        TerminalCore.CopyLineToEntry()
    End Sub

    ' -- Buffer size ----------------------------------------------------------

    Private Sub tlbr_Buffer_Up_Click(sender As Object, e As RoutedEventArgs)
        TerminalCore.StepBufferSize(1)
    End Sub

    Private Sub tlbr_Buffer_Down_Click(sender As Object, e As RoutedEventArgs)
        TerminalCore.StepBufferSize(-1)
    End Sub

    Private Sub txt_BufferSize_PreviewTextInput(sender As Object, e As TextCompositionEventArgs)
        TerminalCore.GuardBufferSizeEntry(e)
    End Sub

    Private Sub txt_BufferSize_PreviewKeyDown(sender As Object, e As KeyEventArgs)
        TerminalCore.HandleBufferSizeKey(e)
    End Sub

    Private Sub txt_BufferSize_LostKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
        TerminalCore.CommitBufferSize()
    End Sub

    ' -- Tool bar -------------------------------------------------------------

    Private Sub tlbr_Terminal_Clear_Click(sender As Object, e As RoutedEventArgs)
        TerminalCore.ClearTranscript()
    End Sub

End Class
