' ============================================================================
'  Controls\DevicePanel.vb
'
'  A WPF custom control that frames one device on the bus.
'
'      +--------------------------------------------------+
'      | DeviceName                                     >  |  header, 85% grey
'      +--------------------------------------------------+
'      |                                                  |  backdrop, 75% grey
'      |   content goes here                              |
'      +--------------------------------------------------+
'
'  The triangle on the right of the header is a disclosure control: hollow orange
'  pointing right when collapsed, filled yellow pointing down when expanded. It
'  shows and hides the backdrop area.
'
'  DevicePanel derives from ContentControl, so a device's registers - a stack of
'  BitFieldEditors, most likely - can simply be written inside it in XAML.
'
'  The control holds a full 256 byte register map. SetRegisterValue and
'  GetRegisterValue are the way in and out of it.
'
'  The default template lives in Themes\Generic.xaml.
' ============================================================================

Imports System.Globalization
Imports System.IO
Imports System.Windows.Controls.Primitives
Imports System.Windows.Documents

<TemplatePart(Name:="PART_Disclosure", Type:=GetType(ToggleButton))>
<TemplatePart(Name:="PART_Content", Type:=GetType(FrameworkElement))>
Public Class DevicePanel
    Inherits ContentControl

    ''' <summary>Register addresses a byte can reach, and therefore the map size.</summary>
    Private Const RegisterCount As Integer = 256

    ''' <summary>
    ''' How far in from the panel's left edge a register's BitFieldEditor starts. Just
    ''' wide enough to read as a vertical bar down the side of the registers, tying
    ''' them together as one device.
    ''' </summary>
    Private Const RegisterIndent As Double = 20


    ''' <summary>
    ''' Raised when LoadDevice could not read the file it was given. Nothing has
    ''' changed by the time it fires; the caller decides how to say so.
    ''' </summary>
    Public Shared ReadOnly LoadFailedEvent As RoutedEvent =
        EventManager.RegisterRoutedEvent("LoadFailed", RoutingStrategy.Bubble,
                                         GetType(RoutedEventHandler), GetType(DevicePanel))

    Public Custom Event LoadFailed As RoutedEventHandler
        AddHandler(handler As RoutedEventHandler)
            Me.AddHandler(LoadFailedEvent, handler)
        End AddHandler
        RemoveHandler(handler As RoutedEventHandler)
            Me.RemoveHandler(LoadFailedEvent, handler)
        End RemoveHandler
        RaiseEvent(sender As Object, e As RoutedEventArgs)
            MyBase.RaiseEvent(e)
        End RaiseEvent
    End Event

    ''' <summary>
    ''' Asked for from the header's right-click menu. The panel does not clone or
    ''' remove itself - it does not own the stack it sits in, so it says what the
    ''' user asked for and lets the workspace do it.
    ''' </summary>
    Public Shared ReadOnly CloneRequestedEvent As RoutedEvent =
        EventManager.RegisterRoutedEvent("CloneRequested", RoutingStrategy.Bubble,
                                         GetType(RoutedEventHandler), GetType(DevicePanel))

    Public Custom Event CloneRequested As RoutedEventHandler
        AddHandler(handler As RoutedEventHandler)
            Me.AddHandler(CloneRequestedEvent, handler)
        End AddHandler
        RemoveHandler(handler As RoutedEventHandler)
            Me.RemoveHandler(CloneRequestedEvent, handler)
        End RemoveHandler
        RaiseEvent(sender As Object, e As RoutedEventArgs)
            MyBase.RaiseEvent(e)
        End RaiseEvent
    End Event

    Public Shared ReadOnly DeleteRequestedEvent As RoutedEvent =
        EventManager.RegisterRoutedEvent("DeleteRequested", RoutingStrategy.Bubble,
                                         GetType(RoutedEventHandler), GetType(DevicePanel))

    Public Custom Event DeleteRequested As RoutedEventHandler
        AddHandler(handler As RoutedEventHandler)
            Me.AddHandler(DeleteRequestedEvent, handler)
        End AddHandler
        RemoveHandler(handler As RoutedEventHandler)
            Me.RemoveHandler(DeleteRequestedEvent, handler)
        End RemoveHandler
        RaiseEvent(sender As Object, e As RoutedEventArgs)
            MyBase.RaiseEvent(e)
        End RaiseEvent
    End Event

    ''' <summary>
    ''' The file this panel was loaded from, or an empty string. Cloning reads it to
    ''' load the same part again.
    ''' </summary>
    Public ReadOnly Property DeviceFile As String
        Get
            Return LoadedFromFile
        End Get
    End Property

    ''' <summary>One byte per register address. Every Byte index is in range.</summary>
    Private ReadOnly RegisterValues(RegisterCount - 1) As Byte

    ''' <summary>Path handed to the last LoadDevice call, or an empty string.</summary>
    Private LoadedFromFile As String = String.Empty

    ''' <summary>
    ''' Height the backdrop keeps when there is nothing loaded, so an empty panel is
    ''' still something to look at. A loaded one is free to shrink to its registers.
    ''' </summary>
    Private Const EmptyBackdropHeight As Double = 40

    ' Template parts.
    Private SummaryBox As TextBlock
    Private TargetBox As ComboBox
    Private TargetHost As FrameworkElement
    Private FunctionBox As TextBox
    Private ContentHost As FrameworkElement

    ''' <summary>
    ''' The editors LoadDevice built, in file order. Kept so collapsing can hide and
    ''' show them without destroying them - the data in them is still wanted.
    ''' </summary>
    Private ReadOnly RegisterEditors As New List(Of BitFieldEditor)

    ''' <summary>Guards the target list while it is being refilled.</summary>
    Private SuppressTargetChange As Boolean = False

    ' What the file said, kept so the target list can be rebuilt whenever the
    ' template turns up. The base address is held separately from DeviceAddress:
    ' picking a target changes DeviceAddress, and the list has to stay anchored to
    ' the address the part actually starts at.
    Private HasDevice As Boolean = False
    Private LoadedBaseAddress As Integer = 0
    Private LoadedAddressBits As Integer = 0


    ' ========================================================================
    '  Dependency properties
    ' ========================================================================

    Public Shared ReadOnly DeviceNameProperty As DependencyProperty =
        DependencyProperty.Register("DeviceName", GetType(String), GetType(DevicePanel),
                                    New FrameworkPropertyMetadata(String.Empty, AddressOf OnHeaderChanged))

    ''' <summary>Caption shown in the header bar.</summary>
    Public Property DeviceName As String
        Get
            Return CStr(GetValue(DeviceNameProperty))
        End Get
        Set(newValue As String)
            SetValue(DeviceNameProperty, newValue)
        End Set
    End Property

    Public Shared ReadOnly DeviceAddressProperty As DependencyProperty =
        DependencyProperty.Register("DeviceAddress", GetType(Byte), GetType(DevicePanel),
                                    New FrameworkPropertyMetadata(CByte(0), AddressOf OnDeviceAddressChanged))

    ''' <summary>The device's address on the bus.</summary>
    Public Property DeviceAddress As Byte
        Get
            Return CByte(GetValue(DeviceAddressProperty))
        End Get
        Set(newValue As Byte)
            SetValue(DeviceAddressProperty, newValue)
        End Set
    End Property

    Public Shared ReadOnly FunctionNameProperty As DependencyProperty =
        DependencyProperty.Register("FunctionName", GetType(String), GetType(DevicePanel),
                                    New FrameworkPropertyMetadata(String.Empty))

    ''' <summary>
    ''' What this device does in this design, as opposed to what part it is. Shown in
    ''' the header on the left, where the user can click it and type.
    ''' </summary>
    Public Property FunctionName As String
        Get
            Return CStr(GetValue(FunctionNameProperty))
        End Get
        Set(newValue As String)
            SetValue(FunctionNameProperty, newValue)
        End Set
    End Property

    Public Shared ReadOnly DeviceDescriptionProperty As DependencyProperty =
        DependencyProperty.Register("DeviceDescription", GetType(String), GetType(DevicePanel),
                                    New FrameworkPropertyMetadata(String.Empty, AddressOf OnHeaderChanged))

    ''' <summary>The part's own description, as the .DEV file gives it.</summary>
    Public Property DeviceDescription As String
        Get
            Return CStr(GetValue(DeviceDescriptionProperty))
        End Get
        Set(newValue As String)
            SetValue(DeviceDescriptionProperty, newValue)
        End Set
    End Property

    Public Shared ReadOnly DevicetypeProperty As DependencyProperty =
        DependencyProperty.Register("Devicetype", GetType(String), GetType(DevicePanel),
                                    New FrameworkPropertyMetadata(String.Empty))

    ''' <summary>What kind of part this is - "EEPROM", "RTC", and so on.</summary>
    Public Property Devicetype As String
        Get
            Return CStr(GetValue(DevicetypeProperty))
        End Get
        Set(newValue As String)
            SetValue(DevicetypeProperty, newValue)
        End Set
    End Property

    Public Shared ReadOnly IsExpandedProperty As DependencyProperty =
        DependencyProperty.Register("IsExpanded", GetType(Boolean), GetType(DevicePanel),
                                    New FrameworkPropertyMetadata(True, AddressOf OnDisclosureChanged))

    ''' <summary>
    ''' State of the disclosure triangle. True points it down, filled and yellow,
    ''' and shows the backdrop; False points it right, hollow and orange, and hides
    ''' the backdrop.
    ''' </summary>
    Public Property IsExpanded As Boolean
        Get
            Return CBool(GetValue(IsExpandedProperty))
        End Get
        Set(newValue As Boolean)
            SetValue(IsExpandedProperty, newValue)
        End Set
    End Property


    ' ========================================================================
    '  Construction
    ' ========================================================================

    Shared Sub New()
        DefaultStyleKeyProperty.OverrideMetadata(GetType(DevicePanel),
                                                 New FrameworkPropertyMetadata(GetType(DevicePanel)))
    End Sub

    Public Overrides Sub OnApplyTemplate()

        MyBase.OnApplyTemplate()

        SummaryBox = TryCast(GetTemplateChild("PART_Summary"), TextBlock)
        TargetHost = TryCast(GetTemplateChild("PART_TargetHost"), FrameworkElement)
        ContentHost = TryCast(GetTemplateChild("bdr_Content"), FrameworkElement)

        If FunctionBox IsNot Nothing Then
            RemoveHandler FunctionBox.PreviewKeyDown, AddressOf FunctionBox_PreviewKeyDown
            RemoveHandler FunctionBox.LostKeyboardFocus, AddressOf FunctionBox_LostKeyboardFocus
        End If

        FunctionBox = TryCast(GetTemplateChild("txt_FunctionName"), TextBox)

        If FunctionBox IsNot Nothing Then
            AddHandler FunctionBox.PreviewKeyDown, AddressOf FunctionBox_PreviewKeyDown
            AddHandler FunctionBox.LostKeyboardFocus, AddressOf FunctionBox_LostKeyboardFocus
        End If

        If TargetBox IsNot Nothing Then
            RemoveHandler TargetBox.SelectionChanged, AddressOf TargetBox_SelectionChanged
        End If

        TargetBox = TryCast(GetTemplateChild("PART_Target"), ComboBox)

        If TargetBox IsNot Nothing Then
            AddHandler TargetBox.SelectionChanged, AddressOf TargetBox_SelectionChanged
        End If

        WireHeaderMenu()

        RefreshHeader()

        ' A device loaded before the template existed never got its target list, so
        ' fill it in now that there is somewhere to put it.
        ApplyTargetAddresses()

        ApplyDisclosure()

    End Sub

    ''' <summary>
    ''' Builds the header's right-click menu and hangs it on the header Border, which
    ''' is what the empty stretch between the name box and Target hit-tests to.
    ''' </summary>
    Private Sub WireHeaderMenu()

        Dim header As FrameworkElement = TryCast(GetTemplateChild("bdr_Header"), FrameworkElement)
        If header Is Nothing Then Exit Sub

        Dim menu As New ContextMenu()

        Dim mnu_Panel_Clone As New MenuItem With {.Header = "_Clone"}
        AddHandler mnu_Panel_Clone.Click, AddressOf Clone_Click
        menu.Items.Add(mnu_Panel_Clone)

        Dim mnu_Panel_Delete As New MenuItem With {.Header = "_Delete"}
        AddHandler mnu_Panel_Delete.Click, AddressOf Delete_Click
        menu.Items.Add(mnu_Panel_Delete)

        header.ContextMenu = menu

    End Sub

    Private Sub Clone_Click(sender As Object, e As RoutedEventArgs)

        MyBase.RaiseEvent(New RoutedEventArgs(CloneRequestedEvent, Me))

    End Sub

    Private Sub Delete_Click(sender As Object, e As RoutedEventArgs)

        MyBase.RaiseEvent(New RoutedEventArgs(DeleteRequestedEvent, Me))

    End Sub


    ' ========================================================================
    '  Collapsing
    '
    '  The triangle does not simply hide the registers. Collapsing keeps the ones
    '  whose padlock is shut - the registers the user has pinned as worth watching -
    '  and hides the rest. Hidden, not discarded: the editors stay in the tree with
    '  their values intact, so nothing is lost by closing a panel.
    '
    '  Which registers survive is decided at the moment the triangle is clicked, and
    '  not re-decided afterwards. If it were, unlocking a register while collapsed
    '  would make it vanish under the pointer with no way to lock it again.
    ' ========================================================================

    Private Shared Sub OnDisclosureChanged(source As DependencyObject, e As DependencyPropertyChangedEventArgs)

        Dim panel As DevicePanel = TryCast(source, DevicePanel)
        If panel IsNot Nothing Then panel.ApplyDisclosure()

    End Sub

    Private Sub ApplyDisclosure()

        If ContentHost Is Nothing Then Exit Sub

        ' A panel holding hand-written content has no registers to sift, so the whole
        ' backdrop follows the triangle as it always did.
        If RegisterEditors.Count = 0 Then
            ContentHost.MinHeight = EmptyBackdropHeight
            ContentHost.Visibility = If(IsExpanded, Visibility.Visible, Visibility.Collapsed)
            Exit Sub
        End If

        ContentHost.MinHeight = 0
        ContentHost.Visibility = Visibility.Visible

        Dim expanded As Boolean = IsExpanded
        Dim seenTop As Boolean = False

        For Each editor As BitFieldEditor In RegisterEditors

            Dim showing As Boolean = expanded OrElse editor.IsLocked

            ' Collapsed rows measure to nothing, so the survivors close up by
            ' themselves and the panel shrinks with them.
            editor.Visibility = If(showing, Visibility.Visible, Visibility.Collapsed)

            If showing Then
                ' Only the topmost row on show draws a top edge; below it each row
                ' borrows the border of the one above.
                editor.BorderThickness = If(seenTop, New Thickness(1, 0, 1, 1), New Thickness(1))
                seenTop = True
            End If

        Next

    End Sub


    ' ========================================================================
    '  Function name
    ' ========================================================================

    ''' <summary>Enter finishes the edit rather than leaving a caret blinking.</summary>
    Private Sub FunctionBox_PreviewKeyDown(sender As Object, e As KeyEventArgs)

        If e.Key <> Key.Enter Then Exit Sub

        CommitFunctionName()
        Keyboard.ClearFocus()
        e.Handled = True

    End Sub

    Private Sub FunctionBox_LostKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)

        CommitFunctionName()

    End Sub

    ''' <summary>
    ''' Clicking anywhere else on the panel ends an edit in progress. WPF does not
    ''' move focus when the click lands on something that cannot take it - a label,
    ''' or bare backdrop - so without this the caret would sit there blinking.
    ''' </summary>
    Protected Overrides Sub OnPreviewMouseDown(e As MouseButtonEventArgs)

        MyBase.OnPreviewMouseDown(e)

        If FunctionBox Is Nothing Then Exit Sub
        If Not FunctionBox.IsKeyboardFocusWithin Then Exit Sub
        If FunctionBox.IsMouseOver Then Exit Sub

        Keyboard.ClearFocus()

    End Sub

    ''' <summary>
    ''' Tidies the typed name once the edit is over: trimmed, and every word
    ''' capitalised the same way register names are.
    ''' </summary>
    Private Sub CommitFunctionName()

        Dim tidied As String = TextTools.CapitaliseWords(If(FunctionName, String.Empty).Trim())

        If Not String.Equals(tidied, FunctionName, StringComparison.Ordinal) Then FunctionName = tidied

    End Sub


    ' ========================================================================
    '  Header summary
    ' ========================================================================

    Private Shared Sub OnHeaderChanged(source As DependencyObject, e As DependencyPropertyChangedEventArgs)

        Dim panel As DevicePanel = TryCast(source, DevicePanel)
        If panel IsNot Nothing Then panel.RefreshHeader()

    End Sub

    ''' <summary>
    ''' The address moved - either from the file or because the user picked another
    ''' target - so the caption and every register's Hostaddress follow it.
    ''' </summary>
    Private Shared Sub OnDeviceAddressChanged(source As DependencyObject, e As DependencyPropertyChangedEventArgs)

        Dim panel As DevicePanel = TryCast(source, DevicePanel)
        If panel Is Nothing Then Exit Sub

        panel.RefreshHeader()
        panel.ApplyHostAddress()

    End Sub

    ''' <summary>
    ''' Tells every register on this panel which device it belongs to. Between
    ''' Hostaddress, Registeraddress and Value, a BitFieldEditor then carries
    ''' everything needed to place a write on the bus.
    ''' </summary>
    Private Sub ApplyHostAddress()

        For Each editor As BitFieldEditor In RegisterEditors
            editor.Hostaddress = DeviceAddress
        Next

    End Sub

    ''' <summary>
    ''' Builds the right hand caption: device name, the address it is answering on,
    ''' and what the part is - each in its own colour.
    ''' </summary>
    Private Sub RefreshHeader()

        If SummaryBox Is Nothing Then Exit Sub

        SummaryBox.Inlines.Clear()

        AddRun(DeviceName, "Brush_Header_Device")
        AddRun(" @ ", "Brush_Header_Plain")
        AddRun(FormattedAddress(), "Brush_Address_Good")
        AddRun(" : ", "Brush_Header_Plain")
        AddRun(DeviceDescription, "Brush_Header_Plain")

    End Sub

    Private Sub AddRun(text As String, brushKey As String)

        SummaryBox.Inlines.Add(New Run(If(text, String.Empty)) With {
            .Foreground = TryCast(TryFindResource(brushKey), Brush)
        })

    End Sub

    ''' <summary>
    ''' The address the panel is set to, as "0101 000 [R/W]". No masking here - this
    ''' is one particular device, not the family, so every bit is known.
    ''' </summary>
    Private Function FormattedAddress() As String

        Return I2CAddress.FormatSevenBit(I2CAddress.SevenBitAddress(CInt(DeviceAddress)))

    End Function


    ' ========================================================================
    '  Register map
    ' ========================================================================

    ''' <summary>
    ''' Stores a value against a register address. Both arguments are bytes, so
    ''' every address from 0 to 255 is valid and nothing can be out of range.
    ''' </summary>
    Public Sub SetRegisterValue(register As Byte, registervalue As Byte)

        RegisterValues(CInt(register)) = registervalue

    End Sub

    ''' <summary>
    ''' Reads back a register. Addresses never written return zero.
    ''' </summary>
    Public Function GetRegisterValue(register As Byte) As Byte

        Return RegisterValues(CInt(register))

    End Function

    ''' <summary>
    ''' Reads a .DEV file and builds the panel from it: the header takes the device's
    ''' name and address, and every register in the file becomes a BitFieldEditor
    ''' stacked down the backdrop.
    '''
    ''' Takes a bare name - "ABCD" finds devices\ABCD.DEV - or a full path.
    '''
    ''' A file that cannot be found or read is not an error worth interrupting for:
    ''' nothing changes and LoadFailed is raised so the caller can put it in the
    ''' status bar. No dialogs.
    ''' </summary>
    Public Sub LoadDevice(devicefile As String)

        Dim resolved As String = ResolveDeviceFile(devicefile)

        If resolved.Length = 0 Then
            MyBase.RaiseEvent(New RoutedEventArgs(LoadFailedEvent, Me))
            Exit Sub
        End If

        Dim data As DeviceFileData

        Try
            data = DeviceLibrary.LoadDeviceFile(resolved)
        Catch
            MyBase.RaiseEvent(New RoutedEventArgs(LoadFailedEvent, Me))
            Exit Sub
        End Try

        LoadedFromFile = resolved

        ApplyDeviceHeader(data)
        ApplyTargetAddresses()
        BuildRegisterEditors(data)

    End Sub

    ''' <summary>
    ''' Turns whatever was passed in into a file that exists, or an empty string.
    ''' Tries it as given first, then as a name in the device library.
    ''' </summary>
    Private Function ResolveDeviceFile(devicefile As String) As String

        If String.IsNullOrWhiteSpace(devicefile) Then Return String.Empty

        Dim candidate As String = devicefile.Trim()

        Try
            If File.Exists(candidate) Then Return candidate

            ' "ABCD" and "ABCD.DEV" both mean devices\ABCD.DEV.
            Dim inLibrary As String = DeviceLibrary.PathFor(Path.GetFileNameWithoutExtension(candidate))
            If File.Exists(inLibrary) Then Return inLibrary

        Catch
            ' A malformed path is just a file that is not there.
        End Try

        Return String.Empty

    End Function

    ''' <summary>Header fields that the .DEV file actually carries.</summary>
    Private Sub ApplyDeviceHeader(data As DeviceFileData)

        DeviceName = If(data.DeviceName, String.Empty).Trim()
        DeviceDescription = If(data.DeviceDescription, String.Empty).Trim()

        Dim baseAddress As Integer

        If Integer.TryParse(If(data.BaseAddress, String.Empty).Trim(), NumberStyles.None,
                            CultureInfo.InvariantCulture, baseAddress) AndAlso
           baseAddress >= 0 AndAlso baseAddress <= Byte.MaxValue Then

            DeviceAddress = CByte(baseAddress)
            LoadedBaseAddress = baseAddress

        Else
            LoadedBaseAddress = CInt(DeviceAddress)
        End If

        Dim bits As Integer

        If Not Integer.TryParse(If(data.AddressBits, String.Empty).Trim(), NumberStyles.None,
                                CultureInfo.InvariantCulture, bits) Then bits = 0

        If bits < 0 Then bits = 0
        If bits > I2CAddress.MaxAddressBits Then bits = I2CAddress.MaxAddressBits

        LoadedAddressBits = bits
        HasDevice = True

        ' The file has no device type field, so Devicetype is left as the caller set it.

    End Sub

    ''' <summary>
    ''' One BitFieldEditor per register, stacked. Each starts RegisterIndent in from
    ''' the panel's left edge and runs to its right edge, so they grow with the panel.
    ''' </summary>
    Private Sub BuildRegisterEditors(data As DeviceFileData)

        ' Two columns: the gutter that holds the target selector, and the registers.
        Dim layout As New Grid()
        layout.ColumnDefinitions.Add(New ColumnDefinition With {.Width = New GridLength(RegisterIndent)})
        layout.ColumnDefinitions.Add(New ColumnDefinition With {.Width = New GridLength(1, GridUnitType.Star)})

        RegisterEditors.Clear()

        Dim rowIndex As Integer = 0

        If data.Registers IsNot Nothing Then

            For Each register As DeviceRegisterData In data.Registers

                If register Is Nothing Then Continue For

                layout.RowDefinitions.Add(New RowDefinition With {.Height = GridLength.Auto})

                Dim editor As New BitFieldEditor With {
                    .Registername = If(register.RegisterName, String.Empty),
                    .Registeraddress = ToRegisterAddress(register.RegisterAddress),
                    .Fieldnames = BitNameList(register),
                    .HorizontalAlignment = HorizontalAlignment.Stretch
                }

                Grid.SetRow(editor, rowIndex)
                Grid.SetColumn(editor, 1)
                layout.Children.Add(editor)
                RegisterEditors.Add(editor)

                rowIndex += 1

            Next

        End If

        Me.Content = layout

        ' The editors were built after DeviceAddress was set, so they have not been
        ' told the host yet.
        ApplyHostAddress()

        ' There is something to look at now. Borders and visibility are ApplyDisclosure's
        ' job, so it gets the last word whether or not the value below changes it.
        IsExpanded = True
        ApplyDisclosure()

    End Sub

    ''' <summary>
    ''' Fills the header's target list with every address this part can be strapped
    ''' to, and selects the one it is currently answering on.
    '''
    ''' Safe to call more than once, and it has to be: LoadDevice usually runs before
    ''' the template has been applied - a panel built in code is not measured until
    ''' after it is added to the tree - so OnApplyTemplate calls it again once the
    ''' parts exist.
    ''' </summary>
    Private Sub ApplyTargetAddresses()

        If TargetBox Is Nothing OrElse Not HasDevice Then Exit Sub

        ' Refilling the list churns the selection, which would otherwise write a
        ' half-built value back into DeviceAddress.
        SuppressTargetChange = True

        TargetBox.ItemsSource = TargetAddresses()
        TargetBox.SelectedItem = CInt(DeviceAddress)
        If TargetBox.SelectedIndex < 0 Then TargetBox.SelectedIndex = 0

        SuppressTargetChange = False

        If TargetHost IsNot Nothing Then TargetHost.Visibility = Visibility.Visible

    End Sub

    ''' <summary>
    ''' Base address, then one more for every value the address pins can take. Stops
    ''' at 255 rather than wrapping.
    ''' </summary>
    Private Function TargetAddresses() As List(Of Integer)

        Dim addresses As New List(Of Integer)

        For offset As Integer = 0 To (1 << LoadedAddressBits) - 1

            Dim address As Integer = LoadedBaseAddress + offset
            If address > Byte.MaxValue Then Exit For

            addresses.Add(address)

        Next

        If addresses.Count = 0 Then addresses.Add(LoadedBaseAddress)

        Return addresses

    End Function

    Private Sub TargetBox_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)

        If SuppressTargetChange Then Exit Sub
        If TargetBox Is Nothing OrElse TargetBox.SelectedItem Is Nothing Then Exit Sub

        ' Drives DeviceAddress, which is what the header caption reads.
        DeviceAddress = CByte(CInt(TargetBox.SelectedItem))

    End Sub

    ''' <summary>
    ''' BitFieldEditor takes its captions most significant bit first, which is the
    ''' order D7 down to D0 is written in.
    ''' </summary>
    Private Shared Function BitNameList(register As DeviceRegisterData) As String

        Return String.Join(",", {register.D7, register.D6, register.D5, register.D4,
                                 register.D3, register.D2, register.D1, register.D0})

    End Function

    ''' <summary>
    ''' Registeraddress is a Byte, so -1 - a device with no register addressing at
    ''' all - cannot be held. It becomes 0, which is the only register such a device
    ''' has anyway.
    ''' </summary>
    Private Shared Function ToRegisterAddress(address As Integer) As Byte

        If address <= 0 Then Return 0
        If address >= Byte.MaxValue Then Return Byte.MaxValue

        Return CByte(address)

    End Function

End Class
