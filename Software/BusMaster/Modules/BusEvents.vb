' ============================================================================
'  Modules\BusEvents.vb
'
'  What happens when a register is changed on screen, and the helpers that go
'  with it.
'
'  There is exactly one handler for every BitFieldEditor in the program. It is
'  attached with EventManager.RegisterClassHandler rather than by hooking each
'  control as it is made: registering against the type covers every instance
'  wherever it lives, including the panels the user clones at run time, so there
'  is no creation site that can be forgotten.
'
'  The handler reads everything it needs off the control that raised the event
'  and off the DevicePanel holding it, gathers it into a RegisterChange, and acts
'  on that. Anything else that needs to put a line on the bus - WriteRegister,
'  ReadRegister, and whatever else follows - belongs in this file, because it
'  belongs to the same event.
' ============================================================================

Imports System.Globalization

''' <summary>
''' Everything known about one register change: what was written, where, and what
''' the device it belongs to is called.
''' </summary>
Public Class RegisterChange

    Public Property Hostaddress As Byte
    Public Property Registeraddress As Byte
    Public Property Value As Byte
    Public Property Registername As String = String.Empty

    Public Property DeviceName As String = String.Empty
    Public Property Devicetype As String = String.Empty
    Public Property FunctionName As String = String.Empty

End Class


Public Module BusEvents

    Private Registered As Boolean = False

    ''' <summary>
    ''' Attaches the one handler that serves every BitFieldEditor. Called once,
    ''' from AppCore.Attach, before any device is loaded.
    ''' </summary>
    Public Sub Initialise()

        If Registered Then Exit Sub

        EventManager.RegisterClassHandler(GetType(BitFieldEditor),
                                          BitFieldEditor.ValueChangedEvent,
                                          New RoutedEventHandler(AddressOf BitFieldValueChanged))

        Registered = True

    End Sub


    ' ========================================================================
    '  The handler
    ' ========================================================================

    ''' <summary>
    ''' A bit was clicked or a value was typed, anywhere in the program.
    ''' </summary>
    Private Sub BitFieldValueChanged(sender As Object, e As RoutedEventArgs)

        Dim editor As BitFieldEditor = TryCast(sender, BitFieldEditor)
        If editor Is Nothing Then Exit Sub

        Dim change As RegisterChange = Describe(editor)

        If AppCore.IsRecording Then
            AddLogentry(change.Hostaddress, change.Registeraddress, change.Value, CommentFor(change))
        End If

    End Sub

    ''' <summary>
    ''' Reads the change off the control and the device panel around it. A control
    ''' with no panel above it - one sitting in a dialog, say - still describes
    ''' itself; the device fields are simply blank.
    ''' </summary>
    Private Function Describe(editor As BitFieldEditor) As RegisterChange

        Dim change As New RegisterChange With {
            .Hostaddress = editor.Hostaddress,
            .Registeraddress = editor.Registeraddress,
            .Value = editor.Value,
            .Registername = If(editor.Registername, String.Empty)
        }

        Dim panel As DevicePanel = OwningPanel(editor)

        If panel IsNot Nothing Then
            change.DeviceName = If(panel.DeviceName, String.Empty)
            change.Devicetype = If(panel.Devicetype, String.Empty)
            change.FunctionName = If(panel.FunctionName, String.Empty)
        End If

        Return change

    End Function

    ''' <summary>The device panel a register belongs to, or Nothing.</summary>
    Private Function OwningPanel(editor As BitFieldEditor) As DevicePanel

        Dim node As DependencyObject = editor

        While node IsNot Nothing

            Dim panel As DevicePanel = TryCast(node, DevicePanel)
            If panel IsNot Nothing Then Return panel

            ' The visual tree is the one that goes up through a control's template.
            ' The logical parent is the fallback for anything not yet rendered.
            Dim parent As DependencyObject = Nothing

            If TypeOf node Is Visual Then parent = VisualTreeHelper.GetParent(node)

            If parent Is Nothing Then parent = LogicalTreeHelper.GetParent(node)

            node = parent

        End While

        Return Nothing

    End Function

    ''' <summary>
    ''' Lead-in characters for the comment column. One character rather than a word,
    ''' so the comment itself starts as far left as possible.
    ''' </summary>
    Public Const WriteLeadIn As String = ">"
    Public Const ReadLeadIn As String = "<"

    ''' <summary>
    ''' The line a write puts in the log's comment column:
    '''
    '''     &gt; PCA9555 : Front_panel_I/O.Output
    '''
    ''' Device, then the function it serves, then the register within it - the dot
    ''' reading as "part of", the way a field of a structure would.
    ''' </summary>
    Private Function CommentFor(change As RegisterChange) As String

        Return WriteLeadIn & " " & change.DeviceName & " : " &
               Unspaced(change.FunctionName) & "." & Unspaced(change.Registername)

    End Function

    ''' <summary>
    ''' Spaces to underscores, so a name reads as one word in the log. Display only -
    ''' the panel and the register keep the name the user typed.
    ''' </summary>
    Private Function Unspaced(text As String) As String

        Return If(text, String.Empty).Replace(" "c, "_"c)

    End Function


    ' ========================================================================
    '  Helpers
    ' ========================================================================

    ''' <summary>
    ''' Puts one transaction on the end of the event log. Addresses and values are
    ''' written in decimal whatever the display is set to, so a saved log always
    ''' reads back the same way.
    ''' </summary>
    Public Sub AddLogentry(hostAddress As Byte, registerAddress As Byte, value As Byte, comment As String)

        EventLogCore.AppendEntry(String.Empty,
                                 hostAddress.ToString(CultureInfo.InvariantCulture),
                                 registerAddress.ToString(CultureInfo.InvariantCulture),
                                 value.ToString(CultureInfo.InvariantCulture),
                                 comment)

    End Sub

End Module
