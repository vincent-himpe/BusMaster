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

    ''' <summary>
    ''' A name as it is written where a space would not do: trimmed, with spaces
    ''' turned into underscores. "Front panel I/O" becomes "Front_panel_I/O".
    '''
    ''' The one definition of what a device or register is called outside its own
    ''' label - the event log writes it this way, and the command window reads it
    ''' back the same way, so what is logged can be typed.
    ''' </summary>
    Public Function Symbolic(text As String) As String

        Return If(text, String.Empty).Trim().Replace(" "c, "_"c)

    End Function

End Module
