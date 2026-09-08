' ============================================================================
'  CommandWindow.xaml.vb
'
'  Event handlers ONLY - see Modules\CommandCore.vb for everything they call.
' ============================================================================

Imports System.ComponentModel

Class CommandWindow

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        CommandCore.Attach(Me)
    End Sub

    Private Sub Window_Closing(sender As Object, e As CancelEventArgs)
        CommandCore.HandleWindowClosing(e)
    End Sub

    Private Sub txt_ActiveCommand_PreviewKeyDown(sender As Object, e As KeyEventArgs)
        CommandCore.HandleKey(e)
    End Sub

    Private Sub txt_ActiveCommand_TextChanged(sender As Object, e As TextChangedEventArgs)
        CommandCore.RefreshCompletions()
    End Sub

    Private Sub txt_ActiveCommand_LostKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
        CommandCore.CloseCompletions()
    End Sub

    Private Sub rtb_CommandHistory_PreviewKeyDown(sender As Object, e As KeyEventArgs)
        CommandCore.HandleHistoryKey(e)
    End Sub

    Private Sub rtb_CommandHistory_SelectionChanged(sender As Object, e As RoutedEventArgs)
        CommandCore.HistoryCaretMoved()
    End Sub

End Class
