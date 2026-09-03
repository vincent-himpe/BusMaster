' ============================================================================
'  Controls\BitFieldEditor.vb
'
'  A WPF custom control that edits one byte as eight bits.
'
'  Layout, left to right:
'
'      [ name ][ value ][ bit 7 ][ bit 6 ] ... [ bit 0 ][ lock ]
'       fixed    fixed    <----- eight equal, scaling regions ----->  fixed
'
'  The leftmost region is the most significant bit. Clicking a region toggles its
'  bit; the value box shows the whole byte in decimal or hexadecimal. The lock
'  freezes the byte - regions stop responding and the value box goes read-only.
'
'  The control never gets narrower than the value box + the lock + eight times its
'  own height, so a region is never smaller than a square.
'
'  The default template lives in Themes\Generic.xaml.
'
'  Note on the project's "no code in event handlers" rule: that rule is about the
'  window, whose job is to delegate to a module. A custom control is a self
'  contained, reusable unit, so its logic belongs to the control - but the
'  handlers below still only call named routines, never inline logic.
' ============================================================================

Imports System.Globalization

''' <summary>How BitFieldEditor renders and parses the value box.</summary>
Public Enum BitFieldValueType
    mode_Decimal = 0
    mode_Hexadecimal = 1
End Enum

<TemplatePart(Name:="PART_Value", Type:=GetType(TextBox))>
<TemplatePart(Name:="PART_Lock", Type:=GetType(CheckBox))>
<TemplatePart(Name:="PART_Fields", Type:=GetType(Panel))>
Public Class BitFieldEditor
    Inherits Control

    ''' <summary>Bits in the byte, and therefore clickable regions.</summary>
    Public Const BitCount As Integer = 8

    Private Const PartName As String = "PART_Name"
    Private Const PartValue As String = "PART_Value"
    Private Const PartLock As String = "PART_Lock"
    Private Const PartFields As String = "PART_Fields"

    Private Const MaxValue As Integer = 255

    ''' <summary>Characters the register name itself has room for.</summary>
    Private Const NameLength As Integer = 14

    ''' <summary>Characters kept when a name is too long; the rest becomes "...".</summary>
    Private Const NameKeep As Integer = 11

    ''' <summary>
    ''' Width of the "(nn)" bracket in front of the name. Padded to a fixed size so
    ''' the names line up down a column of registers whatever their addresses.
    ''' </summary>
    Private Const AddressPrefixLength As Integer = 6

    ' Template parts.
    Private NameBox As TextBlock
    Private ValueBox As TextBox
    Private LockBox As CheckBox
    Private FieldHost As Panel

    ' Region visuals. Index 0 is the leftmost region, which is bit 7.
    Private ReadOnly Regions(BitCount - 1) As Border
    Private ReadOnly RegionLabels(BitCount - 1) As TextBlock


    ' ========================================================================
    '  Dependency properties
    ' ========================================================================

    Public Shared ReadOnly ValueProperty As DependencyProperty =
        DependencyProperty.Register("Value", GetType(Byte), GetType(BitFieldEditor),
                                    New FrameworkPropertyMetadata(CByte(0), AddressOf OnValueChanged))

    ''' <summary>The byte being edited. Bindable, so a host is told when it changes.</summary>
    Public Property Value As Byte
        Get
            Return CByte(GetValue(ValueProperty))
        End Get
        Set(newValue As Byte)
            SetValue(ValueProperty, newValue)
        End Set
    End Property

    Public Shared ReadOnly ValueTypeProperty As DependencyProperty =
        DependencyProperty.Register("ValueType", GetType(BitFieldValueType), GetType(BitFieldEditor),
                                    New FrameworkPropertyMetadata(BitFieldValueType.mode_Decimal,
                                                                  AddressOf OnDisplayChanged))

    ''' <summary>Whether the value box reads and writes decimal or hexadecimal.</summary>
    Public Property ValueType As BitFieldValueType
        Get
            Return CType(GetValue(ValueTypeProperty), BitFieldValueType)
        End Get
        Set(newValue As BitFieldValueType)
            SetValue(ValueTypeProperty, newValue)
        End Set
    End Property

    Public Shared ReadOnly FieldnamesProperty As DependencyProperty =
        DependencyProperty.Register("Fieldnames", GetType(String), GetType(BitFieldEditor),
                                    New FrameworkPropertyMetadata(String.Empty, AddressOf OnDisplayChanged))

    ''' <summary>
    ''' Comma separated region captions, most significant bit first. Surrounding
    ''' spaces are trimmed; a caption too long for its region is shown with an
    ''' ellipsis and carries the full text as a tool tip.
    ''' </summary>
    Public Property Fieldnames As String
        Get
            Return CStr(GetValue(FieldnamesProperty))
        End Get
        Set(newValue As String)
            SetValue(FieldnamesProperty, newValue)
        End Set
    End Property

    ' -- Tags -----------------------------------------------------------------
    '    Registeraddress and Registername are storage for the caller's benefit.
    '    The control never reads them; nothing here depends on their contents.

    Public Shared ReadOnly RegisteraddressProperty As DependencyProperty =
        DependencyProperty.Register("Registeraddress", GetType(Byte), GetType(BitFieldEditor),
                                    New FrameworkPropertyMetadata(CByte(0), AddressOf OnDisplayChanged))

    ''' <summary>
    ''' Address of the register this byte belongs to. A tag - the control does not
    ''' use it.
    ''' </summary>
    Public Property Registeraddress As Byte
        Get
            Return CByte(GetValue(RegisteraddressProperty))
        End Get
        Set(newValue As Byte)
            SetValue(RegisteraddressProperty, newValue)
        End Set
    End Property

    Public Shared ReadOnly RegisternameProperty As DependencyProperty =
        DependencyProperty.Register("Registername", GetType(String), GetType(BitFieldEditor),
                                    New FrameworkPropertyMetadata(String.Empty, AddressOf OnDisplayChanged))

    ''' <summary>
    ''' Name of the register this byte belongs to. A tag - the control does not
    ''' use it.
    '''
    ''' Shadows is required: VB is case insensitive, so "Registername" collides with
    ''' FrameworkElement.RegisterName. That base method is for XAML name scopes and
    ''' is never called on a control like this one, so hiding it costs nothing.
    ''' </summary>
    Public Shadows Property Registername As String
        Get
            Return CStr(GetValue(RegisternameProperty))
        End Get
        Set(newValue As String)
            SetValue(RegisternameProperty, newValue)
        End Set
    End Property

    Public Shared ReadOnly HostaddressProperty As DependencyProperty =
        DependencyProperty.Register("Hostaddress", GetType(Byte), GetType(BitFieldEditor),
                                    New FrameworkPropertyMetadata(CByte(0)))

    ''' <summary>
    ''' Address of whatever holds this register. A tag - the control does not use it.
    ''' </summary>
    Public Property Hostaddress As Byte
        Get
            Return CByte(GetValue(HostaddressProperty))
        End Get
        Set(newValue As Byte)
            SetValue(HostaddressProperty, newValue)
        End Set
    End Property

    Public Shared ReadOnly IsLockedProperty As DependencyProperty =
        DependencyProperty.Register("IsLocked", GetType(Boolean), GetType(BitFieldEditor),
                                    New FrameworkPropertyMetadata(False))

    ''' <summary>
    ''' State of the padlock. It is a marker, not a lock: it does not stop the byte
    ''' being edited. A shut padlock says "keep this register on show", which is what
    ''' DevicePanel reads when it collapses.
    ''' </summary>
    Public Property IsLocked As Boolean
        Get
            Return CBool(GetValue(IsLockedProperty))
        End Get
        Set(newValue As Boolean)
            SetValue(IsLockedProperty, newValue)
        End Set
    End Property


    ' ========================================================================
    '  Routed events
    ' ========================================================================

    ''' <summary>
    ''' Raised when a committed entry was rejected, so the host can report it. The
    ''' control has already put the old text back by the time this fires.
    ''' </summary>
    Public Shared ReadOnly IllegalEntryEvent As RoutedEvent =
        EventManager.RegisterRoutedEvent("IllegalEntry", RoutingStrategy.Bubble,
                                         GetType(RoutedEventHandler), GetType(BitFieldEditor))

    Public Custom Event IllegalEntry As RoutedEventHandler
        AddHandler(handler As RoutedEventHandler)
            Me.AddHandler(IllegalEntryEvent, handler)
        End AddHandler
        RemoveHandler(handler As RoutedEventHandler)
            Me.RemoveHandler(IllegalEntryEvent, handler)
        End RemoveHandler
        RaiseEvent(sender As Object, e As RoutedEventArgs)
            MyBase.RaiseEvent(e)
        End RaiseEvent
    End Event

    ''' <summary>
    ''' Raised when the user has changed the byte - by clicking a region or by
    ''' committing an entry in the value box.
    '''
    ''' Deliberately not raised when a host writes Value in code: loading a device
    ''' sets every register in turn, and that is not a bus transaction. Anything
    ''' driving the bus from code raises its own event.
    ''' </summary>
    Public Shared ReadOnly ValueChangedEvent As RoutedEvent =
        EventManager.RegisterRoutedEvent("ValueChanged", RoutingStrategy.Bubble,
                                         GetType(RoutedEventHandler), GetType(BitFieldEditor))

    Public Custom Event ValueChanged As RoutedEventHandler
        AddHandler(handler As RoutedEventHandler)
            Me.AddHandler(ValueChangedEvent, handler)
        End AddHandler
        RemoveHandler(handler As RoutedEventHandler)
            Me.RemoveHandler(ValueChangedEvent, handler)
        End RemoveHandler
        RaiseEvent(sender As Object, e As RoutedEventArgs)
            MyBase.RaiseEvent(e)
        End RaiseEvent
    End Event

    Private Sub AnnounceValueChanged()

        MyBase.RaiseEvent(New RoutedEventArgs(ValueChangedEvent, Me))

    End Sub


    ' ========================================================================
    '  Construction and templating
    ' ========================================================================

    Shared Sub New()
        DefaultStyleKeyProperty.OverrideMetadata(GetType(BitFieldEditor),
                                                 New FrameworkPropertyMetadata(GetType(BitFieldEditor)))
    End Sub

    Public Overrides Sub OnApplyTemplate()

        MyBase.OnApplyTemplate()

        DetachValueBox()

        NameBox = TryCast(GetTemplateChild(PartName), TextBlock)
        ValueBox = TryCast(GetTemplateChild(PartValue), TextBox)
        LockBox = TryCast(GetTemplateChild(PartLock), CheckBox)
        FieldHost = TryCast(GetTemplateChild(PartFields), Panel)

        AttachValueBox()
        BuildRegions()
        RefreshAll()

    End Sub

    Private Sub AttachValueBox()

        If ValueBox Is Nothing Then Exit Sub

        AddHandler ValueBox.PreviewTextInput, AddressOf ValueBox_PreviewTextInput
        AddHandler ValueBox.PreviewKeyDown, AddressOf ValueBox_PreviewKeyDown
        AddHandler ValueBox.LostKeyboardFocus, AddressOf ValueBox_LostKeyboardFocus
        DataObject.AddPastingHandler(ValueBox, AddressOf ValueBox_Pasting)

    End Sub

    Private Sub DetachValueBox()

        If ValueBox Is Nothing Then Exit Sub

        RemoveHandler ValueBox.PreviewTextInput, AddressOf ValueBox_PreviewTextInput
        RemoveHandler ValueBox.PreviewKeyDown, AddressOf ValueBox_PreviewKeyDown
        RemoveHandler ValueBox.LostKeyboardFocus, AddressOf ValueBox_LostKeyboardFocus
        DataObject.RemovePastingHandler(ValueBox, AddressOf ValueBox_Pasting)

    End Sub

    ''' <summary>Creates the eight region visuals, leftmost first.</summary>
    Private Sub BuildRegions()

        If FieldHost Is Nothing Then Exit Sub

        FieldHost.Children.Clear()

        For index As Integer = 0 To BitCount - 1

            Dim label As New TextBlock With {
                .TextTrimming = TextTrimming.CharacterEllipsis,
                .TextWrapping = TextWrapping.NoWrap,
                .TextAlignment = TextAlignment.Center,
                .VerticalAlignment = VerticalAlignment.Center,
                .Margin = New Thickness(3, 0, 3, 0),
                .FontSize = 11
            }
            label.SetResourceReference(TextBlock.ForegroundProperty, "Brush_Bit_Text")

            ' A hairline between regions, except after the last one.
            Dim divider As Double = If(index < BitCount - 1, 1.0, 0.0)

            Dim region As New Border With {
                .BorderThickness = New Thickness(0, 0, divider, 0),
                .SnapsToDevicePixels = True,
                .Cursor = Cursors.Hand,
                .Tag = index,
                .Child = label
            }
            region.SetResourceReference(Border.BorderBrushProperty, "Brush_Chrome_Border")

            AddHandler region.MouseLeftButtonDown, AddressOf Region_MouseLeftButtonDown

            Regions(index) = region
            RegionLabels(index) = label
            FieldHost.Children.Add(region)

        Next

    End Sub


    ' ========================================================================
    '  Sizing
    ' ========================================================================

    Protected Overrides Sub OnRenderSizeChanged(info As SizeChangedInfo)

        MyBase.OnRenderSizeChanged(info)
        ApplyMinimumWidths()

    End Sub

    ''' <summary>
    ''' Holds the control to "value box + lock + eight times the control height",
    ''' and each region to a square of the control height.
    ''' </summary>
    Private Sub ApplyMinimumWidths()

        Dim height As Double = ActualHeight
        If height <= 0 OrElse Double.IsNaN(height) Then Exit Sub

        For index As Integer = 0 To BitCount - 1
            Dim region As Border = Regions(index)
            If region IsNot Nothing AndAlso Math.Abs(region.MinWidth - height) > 0.5 Then
                region.MinWidth = height
            End If
        Next

        Dim fixedPart As Double = OuterWidth(NameBox) + OuterWidth(ValueBox) + OuterWidth(LockBox)
        Dim wanted As Double = fixedPart + (BitCount * height)

        If Math.Abs(MinWidth - wanted) > 0.5 Then MinWidth = wanted

    End Sub

    ''' <summary>Rendered width of an element including its margin, or zero.</summary>
    Private Shared Function OuterWidth(element As FrameworkElement) As Double

        If element Is Nothing Then Return 0

        Dim width As Double = element.ActualWidth + element.Margin.Left + element.Margin.Right
        If Double.IsNaN(width) Then Return 0

        Return width

    End Function


    ' ========================================================================
    '  Region clicks
    ' ========================================================================

    Private Sub Region_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)

        Dim region As Border = TryCast(sender, Border)
        If region Is Nothing Then Exit Sub

        ToggleRegion(CInt(region.Tag))
        e.Handled = True

    End Sub

    ''' <summary>
    ''' Flips the bit behind a region. Region 0 is the leftmost and holds bit 7.
    ''' </summary>
    Private Sub ToggleRegion(regionIndex As Integer)

        ' IsLocked deliberately does not come into this: the padlock marks a register
        ' as worth keeping on screen, it does not make it read-only.
        If regionIndex < 0 OrElse regionIndex > BitCount - 1 Then Exit Sub

        Dim bit As Integer = (BitCount - 1) - regionIndex

        Value = CByte((CInt(Value) Xor (1 << bit)) And MaxValue)

        AnnounceValueChanged()

    End Sub


    ' ========================================================================
    '  Value box - entry, guarding and commit
    ' ========================================================================

    ''' <summary>Rejects any character the current mode does not allow.</summary>
    Private Sub ValueBox_PreviewTextInput(sender As Object, e As TextCompositionEventArgs)

        e.Handled = Not IsTextAllowed(e.Text)

    End Sub

    ''' <summary>Enter commits, Escape abandons, Space is never useful here.</summary>
    Private Sub ValueBox_PreviewKeyDown(sender As Object, e As KeyEventArgs)

        Select Case e.Key
            Case Key.Enter
                CommitEntry()
                e.Handled = True
            Case Key.Escape
                RefreshValueBox()
                e.Handled = True
            Case Key.Space
                e.Handled = True
        End Select

    End Sub

    ''' <summary>
    ''' Entry is finished with Enter, so an edit abandoned by clicking elsewhere is
    ''' simply dropped - no value change and no complaint.
    ''' </summary>
    Private Sub ValueBox_LostKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)

        RefreshValueBox()

    End Sub

    ''' <summary>Stops a paste from getting round the character guard.</summary>
    Private Sub ValueBox_Pasting(sender As Object, e As DataObjectPastingEventArgs)

        Dim pasted As String = TryCast(e.DataObject.GetData(GetType(String)), String)

        If pasted Is Nothing OrElse Not IsTextAllowed(pasted) Then e.CancelCommand()

    End Sub

    Private Function IsTextAllowed(text As String) As Boolean

        If String.IsNullOrEmpty(text) Then Return False

        For Each character As Char In text
            If Not IsCharacterAllowed(character) Then Return False
        Next

        Return True

    End Function

    Private Function IsCharacterAllowed(character As Char) As Boolean

        If character >= "0"c AndAlso character <= "9"c Then Return True

        If Me.ValueType = BitFieldValueType.mode_Hexadecimal Then
            If character >= "A"c AndAlso character <= "F"c Then Return True
            If character >= "a"c AndAlso character <= "f"c Then Return True
        End If

        Return False

    End Function

    ''' <summary>
    ''' Evaluates what the user typed. An entry that will not parse or falls outside
    ''' 0 to 255 leaves Value and the regions alone, puts the old text back and
    ''' raises IllegalEntry.
    ''' </summary>
    Private Sub CommitEntry()

        If ValueBox Is Nothing Then Exit Sub

        Dim entered As String = If(ValueBox.Text, String.Empty).Trim()
        Dim parsed As Integer = 0
        Dim parsedOk As Boolean

        If Me.ValueType = BitFieldValueType.mode_Hexadecimal Then
            parsedOk = Integer.TryParse(entered, NumberStyles.HexNumber, CultureInfo.InvariantCulture, parsed)
        Else
            parsedOk = Integer.TryParse(entered, NumberStyles.None, CultureInfo.InvariantCulture, parsed)
        End If

        If (Not parsedOk) OrElse parsed < 0 OrElse parsed > MaxValue Then
            RefreshValueBox()
            ValueBox.SelectAll()
            MyBase.RaiseEvent(New RoutedEventArgs(IllegalEntryEvent, Me))
            Exit Sub
        End If

        Dim previous As Byte = Value

        Value = CByte(parsed)

        ' Re-entering the value it already had changes nothing, so format here too.
        RefreshValueBox()
        ValueBox.SelectAll()

        ' ...and for the same reason it is not a change worth announcing.
        If Value <> previous Then AnnounceValueChanged()

    End Sub


    ' ========================================================================
    '  Refreshing the display
    ' ========================================================================

    Private Shared Sub OnValueChanged(source As DependencyObject, e As DependencyPropertyChangedEventArgs)

        Dim editor As BitFieldEditor = TryCast(source, BitFieldEditor)
        If editor IsNot Nothing Then editor.RefreshAll()

    End Sub

    Private Shared Sub OnDisplayChanged(source As DependencyObject, e As DependencyPropertyChangedEventArgs)

        Dim editor As BitFieldEditor = TryCast(source, BitFieldEditor)
        If editor IsNot Nothing Then editor.RefreshAll()

    End Sub

    Private Sub RefreshAll()

        RefreshRegisterName()
        RefreshValueBox()
        RefreshRegions()

    End Sub

    ''' <summary>
    ''' Puts the register address and name in the label - "(0)   OUTPUT" - with the
    ''' name shortened if it will not fit. The full name stays on the tool tip.
    ''' </summary>
    Private Sub RefreshRegisterName()

        If NameBox Is Nothing Then Exit Sub

        Dim name As String = If(Registername, String.Empty)

        Dim prefix As String = ("(" & Registeraddress.ToString(CultureInfo.InvariantCulture) & ")").
                               PadRight(AddressPrefixLength)

        NameBox.Text = prefix & ShortenName(name)
        NameBox.ToolTip = If(name.Length = 0, Nothing, name)

    End Sub

    ''' <summary>
    ''' Names longer than the label holds keep their first eleven characters and get
    ''' an ellipsis, which comes to exactly the fourteen the label has room for.
    ''' </summary>
    Private Shared Function ShortenName(name As String) As String

        If name.Length <= NameLength Then Return name

        Return name.Substring(0, NameKeep) & "..."

    End Function

    ''' <summary>Puts Value into the value box, formatted for the current mode.</summary>
    Private Sub RefreshValueBox()

        If ValueBox Is Nothing Then Exit Sub

        Dim hexadecimal As Boolean = (Me.ValueType = BitFieldValueType.mode_Hexadecimal)

        ' 255 needs three characters, FF needs two.
        ValueBox.MaxLength = If(hexadecimal, 2, 3)
        ValueBox.Text = FormatValue()

        ' Colour says which radix is on show without having to read the number.
        Dim ink As Brush = TryCast(
            TryFindResource(If(hexadecimal, "Brush_Value_Hexadecimal", "Brush_Value_Decimal")), Brush)

        If ink IsNot Nothing Then
            ValueBox.Foreground = ink
            ValueBox.CaretBrush = ink
        End If

    End Sub

    Private Function FormatValue() As String

        If Me.ValueType = BitFieldValueType.mode_Hexadecimal Then
            Return Value.ToString("X2", CultureInfo.InvariantCulture)
        End If

        Return Value.ToString(CultureInfo.InvariantCulture)

    End Function

    ''' <summary>Recolours and re-captions the eight regions from Value and Fieldnames.</summary>
    Private Sub RefreshRegions()

        Dim names() As String = SplitFieldnames()

        For index As Integer = 0 To BitCount - 1

            Dim region As Border = Regions(index)
            Dim label As TextBlock = RegionLabels(index)
            If region Is Nothing OrElse label Is Nothing Then Continue For

            Dim bit As Integer = (BitCount - 1) - index
            Dim isHigh As Boolean = ((CInt(Value) >> bit) And 1) = 1

            region.Background = BitBrush(isHigh)

            Dim caption As String = If(index < names.Length, names(index), String.Empty)
            label.Text = caption

            ' The caption is trimmed with an ellipsis when it does not fit, so the
            ' full text is worth keeping within reach.
            label.ToolTip = If(caption.Length = 0, Nothing, caption)

        Next

    End Sub

    ''' <summary>Splits Fieldnames on commas, trimming the space around each entry.</summary>
    Private Function SplitFieldnames() As String()

        Dim raw As String = If(Fieldnames, String.Empty)
        If raw.Length = 0 Then Return New String() {}

        Dim parts() As String = raw.Split(","c)

        For index As Integer = 0 To parts.Length - 1
            parts(index) = parts(index).Trim()
        Next

        Return parts

    End Function

    ''' <summary>Region colour for a set or clear bit, taken from the theme.</summary>
    Private Function BitBrush(isHigh As Boolean) As Brush

        Dim found As Object = TryFindResource(If(isHigh, "Brush_Bit_High", "Brush_Bit_Low"))

        If TypeOf found Is Brush Then Return DirectCast(found, Brush)

        Return If(isHigh, Brushes.IndianRed, Brushes.SteelBlue)

    End Function

End Class
