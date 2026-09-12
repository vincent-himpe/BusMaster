' ============================================================================
'  Modules\Radix.vb
'
'  Showing a stored number in decimal or hexadecimal, and reading one back.
'
'  Three places let the user switch a column between the two - the registers on
'  the main window, the Event Log's Host, Reg and Value, and the Device Editor's
'  Register Address - and each switch is its own, because they are looked at for
'  different reasons. What they must not have is three ideas of what a hex digit
'  is, so the conversion itself lives here.
'
'  The rule everywhere is that **the stored value stays decimal**, whatever is on
'  screen. The radix is a way of reading a number, not a property of it. That is
'  what keeps a saved log or a device file meaning the same thing whichever way
'  the button happened to be pointing when it was written, and it is why these two
'  routines are a display layer over the real property rather than something that
'  rewrites it.
'
'  Anything that is not a plain non-negative number is handed back untouched:
'  a blank cell, a half-typed entry, the "---" older logs wrote for a read, the
'  "-1" the Device Editor uses for a register with no address yet. None of those
'  has a radix, and mangling them would lose what the user typed.
' ============================================================================

Imports System.Globalization

Public Module Radix

    ''' <summary>
    ''' How a stored value should be written on screen. Hexadecimal comes out in
    ''' upper case and at least two digits, so a column of them lines up and reads
    ''' as bytes.
    ''' </summary>
    Public Function ToDisplay(stored As String, hexadecimal As Boolean) As String

        Dim text As String = If(stored, String.Empty).Trim()

        If Not hexadecimal Then Return text

        Dim number As Integer
        If Not Integer.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, number) Then Return text
        If number < 0 Then Return text

        Return number.ToString("X2", CultureInfo.InvariantCulture)

    End Function

    ''' <summary>
    ''' What to store for something typed while the column was showing that radix.
    ''' The mirror of ToDisplay: what comes out of one goes back in through the
    ''' other and arrives as it started.
    ''' </summary>
    Public Function FromDisplay(typed As String, hexadecimal As Boolean) As String

        Dim text As String = If(typed, String.Empty).Trim()

        If Not hexadecimal Then Return text
        If text.Length = 0 Then Return text

        Dim number As Integer
        If Not Integer.TryParse(text, NumberStyles.HexNumber, CultureInfo.InvariantCulture, number) Then Return text

        Return number.ToString(CultureInfo.InvariantCulture)

    End Function

    ''' <summary>"HEX" or "DEC" - what the button that does the switching reads.</summary>
    Public Function Caption(hexadecimal As Boolean) As String

        Return If(hexadecimal, "HEX", "DEC")

    End Function

    ''' <summary>
    ''' The theme key for the colour that goes with a radix. The same red and blue
    ''' the main window's button uses, so all three switches read alike.
    ''' </summary>
    Public Function BrushKey(hexadecimal As Boolean) As String

        Return If(hexadecimal, "Brush_Value_Hexadecimal", "Brush_Value_Decimal")

    End Function

End Module
