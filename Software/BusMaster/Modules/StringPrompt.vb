' ============================================================================
'  Modules\StringPrompt.vb
'
'  "Type something and press OK" - in this program's colours.
'
'  It replaces Microsoft.VisualBasic.Interaction.InputBox, which draws itself in
'  the system's own light grey and stands out against everything else here.
'
'  Ask returns Nothing when the user cancels, which is not the same as returning
'  an empty string: one means "leave it alone", the other means "they cleared
'  it". Callers that do not care can treat both as nothing to do.
' ============================================================================

Imports System.Windows.Threading

Public Module StringPrompt

    ''' <summary>
    ''' Puts the question, waits for an answer. Nothing back means cancelled.
    ''' </summary>
    Public Function Ask(owner As Window,
                        title As String,
                        prompt As String,
                        Optional startingValue As String = "") As String

        Dim dialog As New StringPromptWindow

        If owner IsNot Nothing Then dialog.Owner = owner

        dialog.Title = If(title, String.Empty)
        dialog.lbl_Prompt.Text = If(prompt, String.Empty)
        dialog.txt_Value.Text = If(startingValue, String.Empty)

        ModalShade.Cover(dialog)

        Try
            If dialog.ShowDialog().GetValueOrDefault() Then Return dialog.txt_Value.Text
        Finally
            ModalShade.Uncover()
        End Try

        Return Nothing

    End Function

    Public Sub Attach(dialog As StringPromptWindow)

        WindowTheme.ApplyDarkTitleBar(dialog)

        ' There is one thing to do here, so the caret starts in the one box, with
        ' whatever was suggested selected and ready to be typed over.
        dialog.Dispatcher.BeginInvoke(DispatcherPriority.Input,
            New Action(Sub()
                           dialog.txt_Value.Focus()
                           dialog.txt_Value.SelectAll()
                       End Sub))

    End Sub

    Public Sub Accept(dialog As StringPromptWindow)

        dialog.DialogResult = True

    End Sub

    Public Sub Abandon(dialog As StringPromptWindow)

        dialog.DialogResult = False

    End Sub

End Module
