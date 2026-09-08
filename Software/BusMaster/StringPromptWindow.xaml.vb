' ============================================================================
'  StringPromptWindow.xaml.vb
'
'  Event handlers ONLY - see Modules\StringPrompt.vb.
' ============================================================================

Class StringPromptWindow

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        StringPrompt.Attach(Me)
    End Sub

    Private Sub btn_Ok_Click(sender As Object, e As RoutedEventArgs)
        StringPrompt.Accept(Me)
    End Sub

    Private Sub btn_Cancel_Click(sender As Object, e As RoutedEventArgs)
        StringPrompt.Abandon(Me)
    End Sub

End Class
