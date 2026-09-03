' ============================================================================
'  Modules\I2CAddress.vb
'
'  I2C addresses get written down two ways and the two get confused constantly.
'
'    - 7 bit form:  the address on its own, 0 to 127.
'    - 8 bit form:  the address shifted up one place, with the read/write flag in
'                   the bottom bit. The write form is therefore always even.
'
'  So 0x50 and 0xA0 are the same part. Anything above 127 here is taken as the
'  8 bit form, and an odd value in that range is simply wrong - it would be the
'  read address, and the bottom bit is not part of the address at all.
'
'  Pure arithmetic, no user interface, so the bus code can share it later.
' ============================================================================

Imports System.Globalization

''' <summary>What Check made of a typed base address.</summary>
Public Enum I2CAddressState
    state_Empty = 0
    state_OutOfRange = 1
    state_OddEightBit = 2
    state_Good = 3
End Enum

Public Module I2CAddress

    Public Const MaxBaseAddress As Integer = 255

    ''' <summary>Above this a base address is read as the 8 bit form.</summary>
    Public Const SevenBitLimit As Integer = 127

    ''' <summary>A 7 bit address cannot have more selectable bits than it has bits.</summary>
    Public Const MaxAddressBits As Integer = 7

    ''' <summary>Stands in for a bit the address pins choose, not the part type.</summary>
    Private Const MaskCharacter As Char = "X"c

    ''' <summary>
    ''' Reads a typed base address and says whether it makes sense.
    ''' </summary>
    Public Function Check(text As String, ByRef value As Integer) As I2CAddressState

        value = 0

        Dim trimmed As String = If(text, String.Empty).Trim()
        If trimmed.Length = 0 Then Return I2CAddressState.state_Empty

        If Not Integer.TryParse(trimmed, NumberStyles.None, CultureInfo.InvariantCulture, value) Then
            Return I2CAddressState.state_OutOfRange
        End If

        If value < 0 OrElse value > MaxBaseAddress Then Return I2CAddressState.state_OutOfRange

        ' In the 8 bit form the bottom bit is the read/write flag, so it must be 0.
        If value > SevenBitLimit AndAlso (value And 1) = 1 Then Return I2CAddressState.state_OddEightBit

        Return I2CAddressState.state_Good

    End Function

    ''' <summary>
    ''' The 7 bit address a base address really means. Over 127 the read/write bit
    ''' is masked off and the rest shifted down one place; at or below 127 the value
    ''' is already a 7 bit address.
    ''' </summary>
    Public Function SevenBitAddress(baseAddress As Integer) As Integer

        If baseAddress > SevenBitLimit Then Return (baseAddress And &HFE) >> 1

        Return baseAddress

    End Function

    ''' <summary>
    ''' Checks a typed address-bits count. Blank is reported separately so the
    ''' caller can decide whether that is allowed.
    ''' </summary>
    Public Function CheckAddressBits(text As String, ByRef value As Integer) As I2CAddressState

        value = 0

        Dim trimmed As String = If(text, String.Empty).Trim()
        If trimmed.Length = 0 Then Return I2CAddressState.state_Empty

        If Not Integer.TryParse(trimmed, NumberStyles.None, CultureInfo.InvariantCulture, value) Then
            Return I2CAddressState.state_OutOfRange
        End If

        If value < 0 OrElse value > MaxAddressBits Then Return I2CAddressState.state_OutOfRange

        Return I2CAddressState.state_Good

    End Function

    ''' <summary>
    ''' Seven bits as "1010 000 [R/W]" - four bits, three bits, then where the
    ''' read/write flag goes.
    '''
    ''' PadLeft rather than taking the right seven characters of a padded string: a
    ''' value that somehow did not fit in seven bits would then show up as too long
    ''' instead of being silently trimmed into a different address.
    ''' </summary>
    Public Function FormatSevenBit(sevenBit As Integer) As String

        Return FormatSevenBit(sevenBit, 0)

    End Function

    ''' <summary>
    ''' As above, but the lowest addressBits bits are shown as X because the address
    ''' pins choose them, not the part type. Two address bits on 80 gives
    ''' "1010 0XX [R/W]"; four on 96 gives "110X XXX [R/W]". Zero masks nothing.
    '''
    ''' Display only - the stored base address keeps whatever was typed.
    ''' </summary>
    Public Function FormatSevenBit(sevenBit As Integer, addressBits As Integer) As String

        Dim bits As String = Convert.ToString(sevenBit, 2).PadLeft(7, "0"c)

        If addressBits > 0 Then

            Dim masked As Integer = Math.Min(addressBits, MaxAddressBits)
            Dim fixedBits As Integer = bits.Length - masked

            bits = bits.Substring(0, fixedBits) & New String(MaskCharacter, masked)

        End If

        Return bits.Substring(0, 4) & " " & bits.Substring(4, 3) & " [R/W]"

    End Function

End Module
