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

        EventManager.RegisterClassHandler(GetType(BitFieldEditor),
                                          BitFieldEditor.WriteRequestedEvent,
                                          New RoutedEventHandler(AddressOf BitFieldWriteRequest))

        EventManager.RegisterClassHandler(GetType(BitFieldEditor),
                                          BitFieldEditor.ReadRequestedEvent,
                                          New RoutedEventHandler(AddressOf BitFieldReadRequest))

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

        ' How much goes out is the Operation setting's business, not this control's.
        If AppCore.OperationMode = BusOperationMode.mode_DeviceOnChange Then

            Dim panel As DevicePanel = OwningPanel(editor)

            If panel IsNot Nothing Then
                WriteDevice(panel)
                Exit Sub
            End If

            ' A register with no device around it has only itself to write.

        End If

        editor.RequestWrite()

    End Sub


    ' ========================================================================
    '  Whole-device and whole-system operations
    '
    '  Fenced off in the log, so a run of registers that was asked for as one
    '  thing reads as one thing.
    '
    '  The labels are laid out by hand and written out in full rather than built
    '  from a pattern - they are all exactly BlockLabelWidth characters, which is
    '  what makes them line up under the grid's fixed-width font, and that is
    '  easier to see and keep true when the text is right here to look at. The
    '  trailing space in "Read " is what keeps it the same width as "Write".
    ' ========================================================================

    ''' <summary>What every block label measures. Checked by the tests, not by eye.</summary>
    Public Const BlockLabelWidth As Integer = 48

    Private Const BlockBeginDeviceWrite As String = "---=========== Begin Device Write ===========---"
    Private Const BlockEndDeviceWrite As String = "--------------  End Device Write  --------------"

    Private Const BlockBeginDeviceRead As String = "---=========== Begin Device Read  ===========---"
    Private Const BlockEndDeviceRead As String = "--------------- End Device Read  ---------------"

    Private Const BlockBeginSystemWrite As String = "============== Begin System Write =============="
    Private Const BlockEndSystemWrite As String = "===------------ End System Write ------------==="

    Private Const BlockBeginSystemRead As String = "============== Begin System Read  =============="
    Private Const BlockEndSystemRead As String = "===------------ End System Read  ------------==="

    ''' <summary>Writes every register on one device, as one labelled block.</summary>
    Public Sub WriteDevice(panel As DevicePanel)

        If panel Is Nothing Then Exit Sub

        AsBlock(BlockBeginDeviceWrite, BlockEndDeviceWrite, Sub() panel.WriteAllRegisters())

    End Sub

    ''' <summary>Reads every register on one device, as one labelled block.</summary>
    Public Sub ReadDevice(panel As DevicePanel)

        If panel Is Nothing Then Exit Sub

        AsBlock(BlockBeginDeviceRead, BlockEndDeviceRead, Sub() panel.ReadAllRegisters())

    End Sub

    ''' <summary>
    ''' Writes every register on every device on screen, as one labelled block.
    ''' One block for the lot rather than a device block inside it for each: the
    ''' user asked for the whole system, and that is what the label says.
    ''' </summary>
    Public Sub WriteSystem()

        AsBlock(BlockBeginSystemWrite, BlockEndSystemWrite,
                Sub()
                    For Each panel As DevicePanel In AppCore.DevicePanels()
                        panel.WriteAllRegisters()
                    Next
                End Sub)

    End Sub

    ''' <summary>Reads every register on every device on screen, as one block.</summary>
    Public Sub ReadSystem()

        AsBlock(BlockBeginSystemRead, BlockEndSystemRead,
                Sub()
                    For Each panel As DevicePanel In AppCore.DevicePanels()
                        panel.ReadAllRegisters()
                    Next
                End Sub)

    End Sub

    ''' <summary>
    ''' Runs an operation with a pair of block labels around it. Whether to mark it
    ''' at all is decided once at the top: a block that opens has to close, whatever
    ''' happens in between.
    ''' </summary>
    Private Sub AsBlock(beginLabel As String, endLabel As String, work As Action)

        Dim marking As Boolean = AppCore.IsRecording

        If marking Then EventLogAddBlock(beginLabel)

        work()

        If marking Then EventLogAddBlock(endLabel)

    End Sub

    ''' <summary>
    ''' A register is to be written out - because it was just changed, because its
    ''' device is being written, or because Write All is walking the workspace.
    ''' Every write comes through here, whatever set it off.
    ''' </summary>
    Private Sub BitFieldWriteRequest(sender As Object, e As RoutedEventArgs)

        Dim editor As BitFieldEditor = TryCast(sender, BitFieldEditor)
        If editor Is Nothing Then Exit Sub

        Dim change As RegisterChange = Describe(editor)

        WriteRegister(change.Hostaddress, change.Registeraddress, change.Value)

        If AppCore.IsRecording Then
            AddLogentry(change.Hostaddress, change.Registeraddress, change.Value,
                        CommentFor(change, WriteLeadIn))
        End If

    End Sub

    ''' <summary>
    ''' A register has been asked for - by right-clicking its bits, or by anything
    ''' in the program calling RequestRead on it.
    '''
    ''' Putting the answer straight into Value is deliberate: a value set in code
    ''' does not raise ValueChanged, so reading a register back does not come out of
    ''' the other end looking like the user had written to it.
    ''' </summary>
    Private Sub BitFieldReadRequest(sender As Object, e As RoutedEventArgs)

        Dim editor As BitFieldEditor = TryCast(sender, BitFieldEditor)
        If editor Is Nothing Then Exit Sub

        Dim change As RegisterChange = Describe(editor)

        editor.Value = ReadRegister(change.Hostaddress, change.Registeraddress)

        ' Logged like a write, but the value column says "---" rather than what came
        ' back. A replayed line is an instruction, not a recording: this one says
        ' "read this register", and what it returns is for then, not for now.
        If AppCore.IsRecording Then
            AddLogentry(change.Hostaddress, change.Registeraddress, ReadPlaceholder,
                        CommentFor(change, ReadLeadIn))
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
    ''' What stands in the value column for a read. A read carries no value out of
    ''' here - it is a request - and anything in that column would read as data the
    ''' user had put on the bus.
    ''' </summary>
    Public Const ReadPlaceholder As String = "---"

    ''' <summary>
    ''' The line a transaction puts in the log's comment column:
    '''
    '''     &gt; PCA9555 : Front_panel_I/O.Output
    '''     &lt; PCA9555 : Front_panel_I/O.Input
    '''
    ''' Device, then the function it serves, then the register within it - the dot
    ''' reading as "part of", the way a field of a structure would. The lead-in says
    ''' which way it went.
    ''' </summary>
    Private Function CommentFor(change As RegisterChange, leadIn As String) As String

        Return leadIn & " " & change.DeviceName & " : " &
               Unspaced(change.FunctionName) & "." & Unspaced(change.Registername)

    End Function

    ''' <summary>
    ''' Spaces to underscores, so a name reads as one word in the log. Display only -
    ''' the panel and the register keep the name the user typed. The command window
    ''' resolves names through the same routine, so anything the log prints can be
    ''' typed straight back in.
    ''' </summary>
    Private Function Unspaced(text As String) As String

        Return TextTools.Symbolic(text)

    End Function


    ' ========================================================================
    '  Helpers
    ' ========================================================================

    ''' <summary>
    ''' Fetches one register from a device.
    '''
    ''' There is no bus yet, so it makes a number up. This is the one place that
    ''' has to change when there is real hardware on the other end - everything
    ''' that reads a register comes through here.
    ''' </summary>
    Public Function ReadRegister(hostAddress As Byte, registerAddress As Byte) As Byte

        Return CByte(Random.Shared.Next(0, 256))

    End Function

    ''' <summary>
    ''' Sends one register out to a device.
    '''
    ''' There is no bus yet, so nothing leaves the program - only the log line the
    ''' caller writes says it happened. The other half of ReadRegister, and the
    ''' other place the hardware plugs in.
    ''' </summary>
    Public Sub WriteRegister(hostAddress As Byte, registerAddress As Byte, value As Byte)

        ' TODO: put the byte on the bus.

    End Sub

    ''' <summary>
    ''' Puts a comment line on the end of the log - marker "/", nothing in the
    ''' address columns, the text in the comment.
    ''' </summary>
    Public Sub EventLogAddComment(text As String)

        EventLogCore.AppendEntry(EventLogMarkers.Comment,
                                 String.Empty, String.Empty, String.Empty, text)

    End Sub

    ''' <summary>
    ''' The same, as a block label - marker "L". Used in pairs to fence off a run of
    ''' transactions that were asked for as one thing.
    ''' </summary>
    Public Sub EventLogAddBlock(text As String)

        EventLogCore.AppendEntry(EventLogMarkers.Label,
                                 String.Empty, String.Empty, String.Empty, text)

    End Sub

    ''' <summary>
    ''' Puts one transaction on the end of the event log. Addresses and values are
    ''' written in decimal whatever the display is set to, so a saved log always
    ''' reads back the same way.
    ''' </summary>
    Public Sub AddLogentry(hostAddress As Byte, registerAddress As Byte, value As String, comment As String)

        EventLogCore.AppendEntry(String.Empty,
                                 hostAddress.ToString(CultureInfo.InvariantCulture),
                                 registerAddress.ToString(CultureInfo.InvariantCulture),
                                 value,
                                 comment)

    End Sub

    ''' <summary>The same, for a transaction that really does carry a byte.</summary>
    Public Sub AddLogentry(hostAddress As Byte, registerAddress As Byte, value As Byte, comment As String)

        AddLogentry(hostAddress, registerAddress, value.ToString(CultureInfo.InvariantCulture), comment)

    End Sub

End Module
