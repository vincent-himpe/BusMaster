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
    ''' Room down the right hand edge for the group box to close in, matching the
    ''' gutter the registers are already indented from on the left.
    ''' </summary>
    Private Const GroupGutterWidth As Double = 10

    ''' <summary>Line weight of the box drawn round a group.</summary>
    Private Const GroupBoxThickness As Double = 2

    Private Const GroupBoxCorner As Double = 4

    ''' <summary>
    ''' Air above and below a group, so its box has somewhere to sit and the group
    ''' reads as one thing set apart from what is around it.
    ''' </summary>
    Private Const GroupSpacing As Double = 3


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

    ''' <summary>
    ''' The user has taken hold of the grab bars and started to move the panel. Like
    ''' Clone and Delete, the panel only says what was asked for: it does not own the
    ''' stack it sits in, so the workspace runs the drag from here on.
    ''' </summary>
    Public Shared ReadOnly ReorderStartedEvent As RoutedEvent =
        EventManager.RegisterRoutedEvent("ReorderStarted", RoutingStrategy.Bubble,
                                         GetType(RoutedEventHandler), GetType(DevicePanel))

    Public Custom Event ReorderStarted As RoutedEventHandler
        AddHandler(handler As RoutedEventHandler)
            Me.AddHandler(ReorderStartedEvent, handler)
        End AddHandler
        RemoveHandler(handler As RoutedEventHandler)
            Me.RemoveHandler(ReorderStartedEvent, handler)
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
    ''' Identity that outlives being moved, renamed, collapsed or reloaded. Made the
    ''' moment the panel exists, written into the project file, and replaced by the
    ''' stored one when a project is opened. Saved views use it to find their way
    ''' back to the right panel.
    '''
    ''' A plain property rather than a dependency property: nothing binds to it and
    ''' nothing draws it, so it would gain nothing from being one.
    ''' </summary>
    Public Property PanelId As String = String.Empty

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

    ''' <summary>
    ''' Which editors draw their own top edge rather than borrowing the border of
    ''' the row above: the ones with a gap above them, at the two ends of a group.
    ''' One entry per editor, same order.
    ''' </summary>
    Private ReadOnly RegisterTopEdge As New List(Of Boolean)

    ''' <summary>The boxes drawn round groups of registers, top to bottom.</summary>
    Private ReadOnly GroupMarks As New List(Of GroupMark)

    ''' <summary>
    ''' Addresses of the registers the device file ticked as QuickViz - the ones
    ''' whoever drew the part up thought were worth looking at. Read once, when the
    ''' file is loaded.
    ''' </summary>
    Private ReadOnly QuickVizRegisters As New List(Of Integer)

    ''' <summary>
    ''' One group box and the registers it encloses. The editors are held rather
    ''' than the row numbers, because whether the box is worth drawing depends on
    ''' whether any of them survived a collapse.
    ''' </summary>
    Private NotInheritable Class GroupMark

        Public ReadOnly Box As Border
        Public ReadOnly Members As New List(Of BitFieldEditor)

        Public Sub New(drawn As Border)
            Box = drawn
        End Sub

        ''' <summary>True while at least one register in the group is on show.</summary>
        Public Function AnyShowing() As Boolean

            For Each editor As BitFieldEditor In Members
                If editor.Visibility = Visibility.Visible Then Return True
            Next

            Return False

        End Function

    End Class

    ''' <summary>Guards the target list while it is being refilled.</summary>
    Private SuppressTargetChange As Boolean = False

    ''' <summary>True while a group's registers are being brought into line.</summary>
    Private SyncingGroups As Boolean = False

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

    Public Sub New()

        ' Every panel has an identity from the moment it exists, however it came to
        ' - the toolbar, a clone, or straight out of the XAML. Opening a project
        ' puts the stored one back in its place.
        PanelId = Guid.NewGuid().ToString()

        ' Read-the-whole-device is settled between this panel and its own registers.
        ' It bubbles up from whichever one was double-clicked, is dealt with here,
        ' and goes no further.
        Me.AddHandler(BitFieldEditor.ReadAllRequestedEvent,
                      New RoutedEventHandler(AddressOf Registers_ReadAllRequested))

        ' So is keeping a group's registers on show together. The panel is the only
        ' thing that knows which registers are one register.
        Me.AddHandler(BitFieldEditor.LockChangedEvent,
                      New RoutedEventHandler(AddressOf Registers_LockChanged))

    End Sub

    ' ========================================================================
    '  Whole-device buttons on the header
    '
    '  The same two operations Ctrl+right-click already offers, put where they
    '  can be found. Both go through BusEvents so the log gets its block labels
    '  and every register still travels the single-read and single-write path.
    ' ========================================================================

    Private WriteAllButton As Button
    Private ReadAllButton As Button

    Private Sub WriteAll_Click(sender As Object, e As RoutedEventArgs)

        BusEvents.WriteDevice(Me)

    End Sub

    Private Sub ReadAll_Click(sender As Object, e As RoutedEventArgs)

        BusEvents.ReadDevice(Me)

    End Sub

    ''' <summary>
    ''' A Ctrl+right-click on any register asks for the whole device.
    '''
    ''' Handed to BusEvents rather than done here, so the log gets its block labels
    ''' around it. The panel knows how to read its registers; it does not know, and
    ''' should not know, that there is a log at all.
    ''' </summary>
    Private Sub Registers_ReadAllRequested(sender As Object, e As RoutedEventArgs)

        BusEvents.ReadDevice(Me)
        e.Handled = True

    End Sub

    ''' <summary>
    ''' Reads every register in this panel, one after another, in the order the
    ''' device file lists them. Each goes out through its own RequestRead, so a
    ''' whole-device read is nothing more than the single reads it is made of.
    ''' </summary>
    Public Sub ReadAllRegisters()

        For Each editor As BitFieldEditor In RegisterEditors
            editor.RequestRead()
        Next

    End Sub

    ''' <summary>
    ''' Writes every register in this panel, in the same order and by the same rule:
    ''' a whole-device write is the single writes it is made of. This is what
    ''' "Device On Change" does after any one register is touched, and what Write
    ''' All does to every panel in turn.
    ''' </summary>
    Public Sub WriteAllRegisters()

        For Each editor As BitFieldEditor In RegisterEditors
            editor.RequestWrite()
        Next

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

        If GrabZone IsNot Nothing Then
            RemoveHandler GrabZone.PreviewMouseLeftButtonDown, AddressOf Grab_PreviewMouseLeftButtonDown
            RemoveHandler GrabZone.PreviewMouseLeftButtonUp, AddressOf Grab_PreviewMouseLeftButtonUp
            RemoveHandler GrabZone.PreviewMouseMove, AddressOf Grab_PreviewMouseMove
            RemoveHandler GrabZone.LostMouseCapture, AddressOf Grab_LostMouseCapture
        End If

        GrabZone = TryCast(GetTemplateChild("PART_Grab"), FrameworkElement)

        If GrabZone IsNot Nothing Then
            AddHandler GrabZone.PreviewMouseLeftButtonDown, AddressOf Grab_PreviewMouseLeftButtonDown
            AddHandler GrabZone.PreviewMouseLeftButtonUp, AddressOf Grab_PreviewMouseLeftButtonUp
            AddHandler GrabZone.PreviewMouseMove, AddressOf Grab_PreviewMouseMove
            AddHandler GrabZone.LostMouseCapture, AddressOf Grab_LostMouseCapture
        End If

        If WriteAllButton IsNot Nothing Then
            RemoveHandler WriteAllButton.Click, AddressOf WriteAll_Click
        End If

        WriteAllButton = TryCast(GetTemplateChild("PART_WriteAll"), Button)

        If WriteAllButton IsNot Nothing Then
            AddHandler WriteAllButton.Click, AddressOf WriteAll_Click
        End If

        If ReadAllButton IsNot Nothing Then
            RemoveHandler ReadAllButton.Click, AddressOf ReadAll_Click
        End If

        ReadAllButton = TryCast(GetTemplateChild("PART_ReadAll"), Button)

        If ReadAllButton IsNot Nothing Then
            AddHandler ReadAllButton.Click, AddressOf ReadAll_Click
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

    ' ========================================================================
    '  The grab bars
    '
    '  Pressing them arms a drag; the drag itself only starts once the pointer has
    '  moved far enough to mean it, so a click that lands here - to dismiss an edit
    '  in the name box, say - still behaves like a click.
    ' ========================================================================

    Private GrabZone As FrameworkElement
    Private GrabOrigin As Point
    Private GrabArmed As Boolean = False

    Private Sub Grab_PreviewMouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)

        GrabOrigin = e.GetPosition(Me)
        GrabArmed = True

        ' Held from the moment the button goes down. Moving far enough to mean a
        ' drag also means leaving the bars, and without the capture the move events
        ' would stop arriving the instant the pointer crossed off them.
        GrabZone.CaptureMouse()

    End Sub

    Private Sub Grab_PreviewMouseLeftButtonUp(sender As Object, e As MouseButtonEventArgs)

        ReleaseGrab()

    End Sub

    Private Sub Grab_LostMouseCapture(sender As Object, e As MouseEventArgs)

        GrabArmed = False

    End Sub

    Private Sub Grab_PreviewMouseMove(sender As Object, e As MouseEventArgs)

        If Not GrabArmed Then Exit Sub

        If e.LeftButton <> MouseButtonState.Pressed Then
            ReleaseGrab()
            Exit Sub
        End If

        Dim moved As Vector = e.GetPosition(Me) - GrabOrigin

        If Math.Abs(moved.X) < SystemParameters.MinimumHorizontalDragDistance AndAlso
           Math.Abs(moved.Y) < SystemParameters.MinimumVerticalDragDistance Then Exit Sub

        ' Handed over. The capture goes back first, because the workspace takes one
        ' of its own the moment it hears about this.
        ReleaseGrab()

        MyBase.RaiseEvent(New RoutedEventArgs(ReorderStartedEvent, Me))

    End Sub

    Private Sub ReleaseGrab()

        GrabArmed = False

        If GrabZone IsNot Nothing AndAlso GrabZone.IsMouseCaptured Then GrabZone.ReleaseMouseCapture()

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

        For index As Integer = 0 To RegisterEditors.Count - 1

            Dim editor As BitFieldEditor = RegisterEditors(index)
            Dim showing As Boolean = expanded OrElse editor.IsLocked

            ' Collapsed rows measure to nothing, so the survivors close up by
            ' themselves and the panel shrinks with them.
            editor.Visibility = If(showing, Visibility.Visible, Visibility.Collapsed)

            If showing Then
                ' Only the topmost row on show draws a top edge; below it each row
                ' borrows the border of the one above. A row at the edge of a group
                ' is the exception: there is a gap above it, so there is no border
                ' up there to borrow.
                Dim ownTop As Boolean = Not seenTop OrElse RegisterTopEdge(index)

                editor.BorderThickness = If(ownTop, New Thickness(1), New Thickness(1, 0, 1, 1))
                seenTop = True
            End If

        Next

        ' A box with nothing left inside it would draw as a thin sliver, so it goes
        ' when the last of its registers does.
        For Each mark As GroupMark In GroupMarks
            mark.Box.Visibility = If(mark.AnyShowing(), Visibility.Visible, Visibility.Collapsed)
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
    ''' The register editors this panel built, in the order the device file lists
    ''' them. A copy, so a caller walking them cannot disturb the panel's own list.
    ''' </summary>
    Public Function Registers() As List(Of BitFieldEditor)

        Return New List(Of BitFieldEditor)(RegisterEditors)

    End Function

    ''' <summary>
    ''' Addresses of the registers whose padlock is shut. This is what a saved view
    ''' records - the padlock is what decides which registers survive a collapse.
    ''' </summary>
    Public Function GetLockedRegisters() As List(Of Integer)

        Dim locked As New List(Of Integer)

        For Each editor As BitFieldEditor In RegisterEditors
            If editor.IsLocked Then locked.Add(CInt(editor.Registeraddress))
        Next

        Return locked

    End Function

    ''' <summary>
    ''' Shuts the padlock on the listed registers and opens every other one, so a
    ''' view restores what was saved rather than adding to whatever is set now.
    ''' Addresses this device does not have are ignored, which is what lets a view
    ''' survive the device file changing underneath it.
    ''' </summary>
    Public Sub SetLockedRegisters(locked As IEnumerable(Of Integer))

        Dim wanted As New HashSet(Of Integer)

        If locked IsNot Nothing Then
            For Each address As Integer In locked
                wanted.Add(address)
            Next
        End If

        ' Applied without the group rule getting in the way, so the result does not
        ' depend on which register of a group the list happens to name first.
        SyncingGroups = True

        For Each editor As BitFieldEditor In RegisterEditors
            editor.IsLocked = wanted.Contains(CInt(editor.Registeraddress))
        Next

        SyncingGroups = False

        NormaliseGroupVisibility()

        ' A collapsed panel is showing the locked registers right now, so what it
        ' shows has just changed.
        ApplyDisclosure()

    End Sub

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
    ''' How many of this device's registers are ticked for QuickViz.
    ''' </summary>
    Public Function QuickVizCount() As Integer

        Return QuickVizRegisters.Count

    End Function

    ''' <summary>
    ''' Shows exactly the registers the device file ticked for QuickViz and hides
    ''' the rest, so a part with fifty configuration registers and three data ones
    ''' collapses to the three that are worth watching.
    '''
    ''' It is the same thing a saved view does, with the list coming from the device
    ''' file rather than from the project - which is what lets it work on a device
    ''' the moment it is added, with nothing saved anywhere.
    ''' </summary>
    Public Sub ApplyQuickView()

        SetLockedRegisters(QuickVizRegisters)

    End Sub

    ''' <summary>
    ''' Reads this panel's device file again, for when the file has been edited
    ''' underneath it. The register list is rebuilt from the file, and everything
    ''' the user had set up on the panel is put back on top of it:
    '''
    '''   - the target address, if the part can still be strapped to it
    '''   - which registers have their visibility switch on, matched by address
    '''   - the byte showing in each register, matched by address
    '''   - open or closed
    '''
    ''' The function name is not touched - LoadDevice never writes it.
    '''
    ''' Values are put back by address, which is what "the same register" can mean
    ''' without reading the bus again. A register the edit removed takes its value
    ''' with it, and one it added starts at zero. Nothing here is a bus operation:
    ''' assigning Value does not raise ValueChanged, so no write goes out and no
    ''' line lands in the log.
    '''
    ''' Does nothing for a panel with no file behind it.
    ''' </summary>
    Public Sub ReloadDevice()

        If LoadedFromFile.Length = 0 Then Exit Sub

        Dim wasExpanded As Boolean = IsExpanded
        Dim wasTargeting As Integer = CInt(DeviceAddress)
        Dim wasLocked As List(Of Integer) = GetLockedRegisters()

        Dim wasShowing As New Dictionary(Of Byte, Byte)

        For Each editor As BitFieldEditor In RegisterEditors
            wasShowing(editor.Registeraddress) = editor.Value
        Next

        LoadDevice(LoadedFromFile)

        ' Only if the part still answers there - the edit may have moved its base
        ' address or changed how many pins choose it, and then the old target means
        ' nothing. ApplyTargetAddresses is what puts the header and the drop-down
        ' back in step with each other.
        If TargetAddresses().Contains(wasTargeting) Then
            DeviceAddress = CByte(wasTargeting)
            ApplyTargetAddresses()
        End If

        For Each editor As BitFieldEditor In RegisterEditors

            Dim showing As Byte

            If wasShowing.TryGetValue(editor.Registeraddress, showing) Then editor.Value = showing

        Next

        SetLockedRegisters(wasLocked)

        ' Last word, because BuildRegisterEditors opens the panel to show that it
        ' has something in it.
        IsExpanded = wasExpanded

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
    ''' the panel's left edge and stops GroupGutterWidth short of its right one, so
    ''' they grow with the panel and leave the gutters a group box can be drawn in.
    ''' </summary>
    Private Sub BuildRegisterEditors(data As DeviceFileData)

        ' Three columns: the gutter that holds the target selector, the registers,
        ' and the gutter the group boxes close in.
        Dim layout As New Grid()
        layout.ColumnDefinitions.Add(New ColumnDefinition With {.Width = New GridLength(RegisterIndent)})
        layout.ColumnDefinitions.Add(New ColumnDefinition With {.Width = New GridLength(1, GridUnitType.Star)})
        layout.ColumnDefinitions.Add(New ColumnDefinition With {.Width = New GridLength(GroupGutterWidth)})

        RegisterEditors.Clear()
        RegisterTopEdge.Clear()
        GroupMarks.Clear()
        QuickVizRegisters.Clear()

        Dim tags As New List(Of String)
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

                ' Kept by the address the editor ended up with, not the one the file
                ' wrote, so the two cannot disagree about which register this is.
                If register.QuickViz Then QuickVizRegisters.Add(CInt(editor.Registeraddress))

                Grid.SetRow(editor, rowIndex)
                Grid.SetColumn(editor, 1)
                layout.Children.Add(editor)
                RegisterEditors.Add(editor)
                RegisterTopEdge.Add(False)
                tags.Add(If(register.RegisterGroup, String.Empty).Trim())

                rowIndex += 1

            Next

        End If

        ' Added last but given a lower z-index, so the fill sits behind the registers
        ' while the code that needs the editors in hand still runs after them.
        BuildGroupMarks(layout, tags)

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
    ''' A register's visibility switch was thrown. Registers in a group are one
    ''' register spread over several addresses, so they are shown and hidden
    ''' together: half a 16 bit register surviving a collapse would be nothing worth
    ''' looking at.
    ''' </summary>
    Private Sub Registers_LockChanged(sender As Object, e As RoutedEventArgs)

        e.Handled = True
        MatchGroupVisibility(TryCast(e.OriginalSource, BitFieldEditor))

    End Sub

    ''' <summary>
    ''' Puts the rest of a register's group where that register has just gone. Does
    ''' nothing for a register that is in no group.
    ''' </summary>
    Private Sub MatchGroupVisibility(changed As BitFieldEditor)

        ' Setting the others throws their switches too, and those come back through
        ' here. One pass is enough.
        If changed Is Nothing OrElse SyncingGroups Then Exit Sub

        Dim mark As GroupMark = MarkHolding(changed)
        If mark Is Nothing Then Exit Sub

        SyncingGroups = True

        For Each member As BitFieldEditor In mark.Members
            member.IsLocked = changed.IsLocked
        Next

        SyncingGroups = False

        ' What a collapsed panel is showing has just changed.
        ApplyDisclosure()

    End Sub

    ''' <summary>
    ''' Brings every group into line with itself: a group with any register on show
    ''' shows all of them. For after something has set the switches wholesale - a
    ''' restored view knows nothing about groups, and one written before groups
    ''' existed can name half of one.
    ''' </summary>
    Private Sub NormaliseGroupVisibility()

        SyncingGroups = True

        For Each mark As GroupMark In GroupMarks

            Dim anyShown As Boolean = False

            For Each member As BitFieldEditor In mark.Members
                If member.IsLocked Then anyShown = True
            Next

            For Each member As BitFieldEditor In mark.Members
                member.IsLocked = anyShown
            Next

        Next

        SyncingGroups = False

    End Sub

    ''' <summary>The group this register belongs to, or Nothing if it is on its own.</summary>
    Private Function MarkHolding(editor As BitFieldEditor) As GroupMark

        For Each mark As GroupMark In GroupMarks
            If mark.Members.Contains(editor) Then Return mark
        Next

        Return Nothing

    End Function

    ''' <summary>
    ''' Draws a box round each run of registers sharing a group tag - one register
    ''' spread over consecutive addresses, so it should read as one thing.
    '''
    ''' The box spans all three columns, so it closes in the gutters either side of
    ''' the registers rather than crossing them. Neighbouring groups take different
    ''' colours, alternating down the panel, because two boxes touching end to end
    ''' would otherwise read as one. An untagged register gets nothing.
    ''' </summary>
    Private Sub BuildGroupMarks(layout As Grid, tags As List(Of String))

        Dim first As Integer = 0

        While first < tags.Count

            Dim tag As String = tags(first)

            If tag.Length = 0 Then
                first += 1
                Continue While
            End If

            ' How far the run of this tag reaches.
            Dim last As Integer = first

            While last + 1 < tags.Count AndAlso
                  String.Equals(tags(last + 1), tag, StringComparison.OrdinalIgnoreCase)
                last += 1
            End While

            AddGroupMark(layout, first, last)

            first = last + 1

        End While

    End Sub

    ''' <summary>
    ''' Boxes the registers from first to last, and opens up the space above and
    ''' below that the box is drawn in.
    ''' </summary>
    Private Sub AddGroupMark(layout As Grid, first As Integer, last As Integer)

        ' Each end of the group is pushed away from whatever is next to it. Two
        ' groups meeting therefore get both gaps, which is what two boxes need.
        RegisterEditors(first).Margin = New Thickness(0, GroupSpacing, 0, RegisterEditors(first).Margin.Bottom)
        RegisterEditors(last).Margin = New Thickness(0, RegisterEditors(last).Margin.Top, 0, GroupSpacing)

        ' A row with a gap above it has no neighbour to borrow a top border from.
        RegisterTopEdge(first) = True
        If last + 1 < RegisterEditors.Count Then RegisterTopEdge(last + 1) = True

        Dim brushKey As String = If(GroupMarks.Count Mod 2 = 0, "Brush_Group_Box_1", "Brush_Group_Box_2")
        Dim ink As Brush = TryCast(TryFindResource(brushKey), Brush)

        ' Filled, not just outlined. The registers are opaque and sit on top, so all
        ' that shows of the fill is the part of the balloon they do not cover: a
        ' solid bar down the gutter either side of them.
        Dim box As New Border With {
            .BorderBrush = ink,
            .Background = ink,
            .BorderThickness = New Thickness(GroupBoxThickness),
            .CornerRadius = New CornerRadius(GroupBoxCorner),
            .Margin = New Thickness(2, 1, 2, 1),
            .SnapsToDevicePixels = True,
            .IsHitTestVisible = False
        }

        ' Behind the registers. Without this the box is the last child of the grid
        ' and would paint its fill straight over them.
        Panel.SetZIndex(box, -1)

        Grid.SetRow(box, first)
        Grid.SetRowSpan(box, last - first + 1)
        Grid.SetColumn(box, 0)
        Grid.SetColumnSpan(box, layout.ColumnDefinitions.Count)

        layout.Children.Add(box)

        Dim mark As New GroupMark(box)

        For index As Integer = first To last

            mark.Members.Add(RegisterEditors(index))

            ' One switch for the group, on the register at the top of it. The rest
            ' follow it, so a switch of their own could only ever agree with it.
            RegisterEditors(index).ShowVisibilitySwitch = (index = first)

        Next

        GroupMarks.Add(mark)

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
