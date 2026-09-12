' ============================================================================
'  EventLogWindow.xaml.vb
'
'  Event handlers ONLY. Everything else is in Modules\EventLogCore.vb.
'
'  The window is created once at start-up and hidden; closing it hides it again
'  rather than destroying it, so the log survives being put away.
' ============================================================================

Imports System.ComponentModel

Class EventLogWindow

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        EventLogCore.Attach(Me)
    End Sub

    Private Sub Window_Closing(sender As Object, e As CancelEventArgs)
        EventLogCore.HandleWindowClosing(e)
    End Sub

    ' -- Tool bar -------------------------------------------------------------

    Private Sub tlbr_Log_Save_As_Click(sender As Object, e As RoutedEventArgs)
        EventLogCore.SaveLogAs()
    End Sub

    Private Sub tlbr_Log_Save_Click(sender As Object, e As RoutedEventArgs)
        EventLogCore.SaveLog()
    End Sub

    Private Sub tlbr_Log_Replay_Click(sender As Object, e As RoutedEventArgs)
        EventLogCore.ReplayLine()
    End Sub

    Private Sub tlbr_Log_Replay_Block_Click(sender As Object, e As RoutedEventArgs)
        EventLogCore.ReplayBlock()
    End Sub

    Private Sub tlbr_Log_Replay_All_Click(sender As Object, e As RoutedEventArgs)
        EventLogCore.ReplayAll()
    End Sub

    Private Sub tlbr_Log_Clear_Click(sender As Object, e As RoutedEventArgs)
        EventLogCore.ClearLog()
    End Sub

    Private Sub tlbr_Log_ValueType_Click(sender As Object, e As RoutedEventArgs)
        EventLogCore.ToggleRadix()
    End Sub

End Class
