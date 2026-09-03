' ============================================================================
'  Modules\TextTools.vb
'
'  Small text helpers shared by more than one place.
' ============================================================================

Public Module TextTools

    ''' <summary>
    ''' Upper cases the first letter of every word and leaves the rest of each word
    ''' alone, so "open drain" becomes "Open Drain" while "IODIR" stays "IODIR".
    '''
    ''' TextInfo.ToTitleCase is deliberately not used: it lower cases the tail of a
    ''' word, which turns an acronym like IODIR into Iodir.
    ''' </summary>
    Public Function CapitaliseWords(text As String) As String

        If String.IsNullOrEmpty(text) Then Return If(text, String.Empty)

        Dim letters() As Char = text.ToCharArray()
        Dim atWordStart As Boolean = True

        For index As Integer = 0 To letters.Length - 1

            If Char.IsWhiteSpace(letters(index)) Then
                atWordStart = True
            Else
                If atWordStart Then letters(index) = Char.ToUpperInvariant(letters(index))
                atWordStart = False
            End If

        Next

        Return New String(letters)

    End Function

End Module
