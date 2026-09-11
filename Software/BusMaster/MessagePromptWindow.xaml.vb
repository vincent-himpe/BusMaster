' ============================================================================
'  MessagePromptWindow.xaml.vb
'
'  Event handlers ONLY, and the two values the box has to carry from the click
'  back to the caller. See Modules\MessagePrompt.vb.
' ============================================================================

Class MessagePromptWindow

    ''' <summary>Which button was pressed. Read once ShowDialog has returned.</summary>
    Public Property Answer As MessageBoxResult = MessageBoxResult.None

    ''' <summary>What Escape means in this mode - Cancel, No, or OK.</summary>
    Public Property EscapeAnswer As MessageBoxResult = MessageBoxResult.OK

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        MessagePrompt.Attach(Me)
    End Sub

    Private Sub Window_PreviewKeyDown(sender As Object, e As KeyEventArgs)
        MessagePrompt.HandleKey(Me, e)
    End Sub

    Private Sub btn_Ok_Click(sender As Object, e As RoutedEventArgs)
        MessagePrompt.Answered(Me, MessageBoxResult.OK)
    End Sub

    Private Sub btn_Yes_Click(sender As Object, e As RoutedEventArgs)
        MessagePrompt.Answered(Me, MessageBoxResult.Yes)
    End Sub

    Private Sub btn_No_Click(sender As Object, e As RoutedEventArgs)
        MessagePrompt.Answered(Me, MessageBoxResult.No)
    End Sub

    Private Sub btn_Cancel_Click(sender As Object, e As RoutedEventArgs)
        MessagePrompt.Answered(Me, MessageBoxResult.Cancel)
    End Sub

End Class
