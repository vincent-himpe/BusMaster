' ============================================================================
'  Modules\ProbeControl.vb
'
'  The probe on the other end of the USB serial link: what goes out to it, and
'  what comes back.
'
'  Two routines, and every byte passes through one of them:
'
'      ProbeSend      a line on its way out to the probe
'      ProbeReceive   a line that has arrived from it
'
'  Both are deliberately thin for now. Whatever has to be done to a line - a
'  checksum, a framing character, a reply matched to its request - belongs inside
'  these two and nowhere else, the same way every register operation goes through
'  ReadRegister and WriteRegister.
'
'  The port itself is named in settings.json, not in the project: which COM port
'  a probe turns up on belongs to the machine. Leave ProbePort blank and the
'  link turns lines straight round instead, so the Terminal is usable with no
'  hardware attached.
' ============================================================================

Imports System.Globalization
Imports System.IO.Ports
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows.Threading

''' <summary>
''' One port that could hold a probe, and how to write it in a list.
'''
''' A probe says "BusMaster" over the bus, and may say more than that. The common
''' part is dropped from the label - every entry in the list is a BusMaster, so
''' repeating it says nothing - and what is left tells the user which one it is:
'''
'''     iProduct "BusMaster"            -&gt; "(COM7)"
'''     iProduct "BusMaster Myprobe"    -&gt; "Myprobe (COM9)"
'''     no product string at all        -&gt; "COM6"
''' </summary>
Public Class ProbePort

    Public ReadOnly PortName As String
    Public ReadOnly Product As String

    Public Sub New(portName As String, product As String)

        Me.PortName = If(portName, String.Empty)
        Me.Product = If(product, String.Empty)

    End Sub

    Public ReadOnly Property Display As String
        Get
            If PortName.Length = 0 Then Return String.Empty

            If Not Product.StartsWith(ProbeControl.ProbeProductPrefix,
                                      StringComparison.OrdinalIgnoreCase) Then Return PortName

            Dim extra As String = Product.Substring(ProbeControl.ProbeProductPrefix.Length).Trim()

            If extra.Length = 0 Then Return "(" & PortName & ")"

            Return extra & " (" & PortName & ")"
        End Get
    End Property

    ''' <summary>
    ''' What a list shows. Overridden rather than left to DisplayMemberPath: that
    ''' reaches the drop-down but not the closed box, which ends up showing the
    ''' type name.
    ''' </summary>
    Public Overrides Function ToString() As String

        Return Display

    End Function

End Class


Public Module ProbeControl

    ''' <summary>What ends a line on the wire.</summary>
    Public Const LineEnding As String = vbCrLf

    Private Port As SerialPort

    ''' <summary>
    ''' Bytes that have arrived but do not make a whole line yet. Only ever touched
    ''' on the UI thread - see Port_DataReceived.
    ''' </summary>
    Private ReadOnly Arriving As New StringBuilder()

    ''' <summary>
    ''' Set once a port has been reported broken, so a failing link complains once
    ''' rather than on every line.
    ''' </summary>
    Private Complained As Boolean = False


    ' ========================================================================
    '  Out
    ' ========================================================================

    ''' <summary>
    ''' Sends a line to the probe. For now it goes out exactly as it was handed
    ''' over, with a carriage return and line feed on the end.
    ''' </summary>
    Public Sub ProbeSend(text As String)

        ' TODO: whatever the line needs doing to it before it goes out.

        Transmit(If(text, String.Empty) & LineEnding)

    End Sub


    ' ========================================================================
    '  In
    ' ========================================================================

    ''' <summary>
    ''' Takes a line that has arrived from the probe. Splitting the incoming bytes
    ''' into lines is the port's business; making sense of a line is this routine's.
    '''
    ''' Always called on the UI thread - the port hands its data over on a worker
    ''' thread and that is marshalled before it gets this far.
    ''' </summary>
    Public Sub ProbeReceive(text As String)

        ' TODO: whatever has to be made of the line that came back.

        TerminalCore.ShowReceived(If(text, String.Empty))

    End Sub


    ' ========================================================================
    '  Lines the user drives by hand
    '
    '  The four signals and the target's reset line are not bus traffic, so they
    '  do not go through BusEvents - they are wires on the probe, and this is the
    '  probe's file. Each one reports what it was asked to do and has a TODO where
    '  the command to the probe goes; nothing is sent to the hardware yet.
    ' ========================================================================

    ''' <summary>
    ''' One of the four user signals was switched. Source is 1 to 4 - red, green,
    ''' blue, yellow, the order they sit in on the toolbar - and NewState is what
    ''' the switch was just set to.
    ''' </summary>
    Public Sub UserSignal(Source As Integer, NewState As Boolean)

        ' TODO: tell the probe to drive user line Source to NewState.

        AppCore.SetStatus("User signal " & Source.ToString(CultureInfo.InvariantCulture) &
                          If(NewState, " high", " low"))

    End Sub

    ''' <summary>
    ''' The target's reset line was held or released. True holds the part in reset;
    ''' False lets it run.
    ''' </summary>
    Public Sub TargetReset(NewState As Boolean)

        ' TODO: tell the probe to assert or release the target's reset line.

        AppCore.SetStatus(If(NewState, "Target held in reset", "Target released - running"))

    End Sub

    ''' <summary>
    ''' Reset the target and let go again, in one go - the button does not stay
    ''' down, so there is no state to pass.
    ''' </summary>
    Public Sub TargetResetPulse()

        ' TODO: tell the probe to pulse the target's reset line.

        AppCore.SetStatus("Reset pulse sent to the target")

    End Sub

    ''' <summary>
    ''' Let go of the bus: stop driving SCL and SDA and leave both lines to their
    ''' pull-ups, so something else can have a turn.
    ''' </summary>
    Public Sub ReleaseBus()

        ' TODO: tell the probe to float SCL and SDA.

        AppCore.SetStatus("Bus released")

    End Sub

    ''' <summary>
    ''' Clear a bus a slave is holding down: clock SCL until the slave lets SDA go,
    ''' then issue a stop so everything is back at idle.
    '''
    ''' The Bus menu's Reset Bus is still a stub and describes the same operation.
    ''' One of the two should end up calling the other rather than both growing
    ''' their own version.
    ''' </summary>
    Public Sub ClearStuckBus()

        ' TODO: tell the probe to clock out a stuck slave, then send a stop.

        AppCore.SetStatus("Cleared the bus")

    End Sub

    ''' <summary>
    ''' How hard the probe drives SCL and SDA. True is push-pull - both levels
    ''' driven, 0/1 - and False is open drain, 0/Z, which is what I2C is and where
    ''' the switch starts.
    ''' </summary>
    Public Sub SetDriveStrength(NewState As Boolean)

        ' TODO: tell the probe which way to drive the two lines.

        AppCore.SetStatus(If(NewState, "Bus lines driven push-pull (0/1)",
                                       "Bus lines open drain (0/Z)"))

    End Sub


    ' ========================================================================
    '  The port
    ' ========================================================================

    ''' <summary>Whether there is a port open to the probe right now.</summary>
    Public ReadOnly Property IsConnected As Boolean
        Get
            Return Port IsNot Nothing AndAlso Port.IsOpen
        End Get
    End Property

    ''' <summary>The name the settings give for the probe's port, trimmed.</summary>
    Private ReadOnly Property WantedPort As String
        Get
            Return If(AppSettings.Current.ProbePort, String.Empty).Trim()
        End Get
    End Property

    ''' <summary>
    ''' The port that is open right now, or an empty string. Not the same as the one
    ''' the settings name - that is the one to use next.
    ''' </summary>
    Public ReadOnly Property PortInUse As String
        Get
            If Not IsConnected Then Return String.Empty

            Return Port.PortName
        End Get
    End Property

    ''' <summary>
    ''' Tells whatever shows the state of the link that it has changed. One place,
    ''' so a connection made from the toolbar and one made by the first line going
    ''' out both end up on screen the same way.
    ''' </summary>
    Private Sub LinkChanged()

        OnUiThread(Sub()
                       AppCore.RefreshProbeConnection()
                       TerminalCore.ShowProbeState()
                   End Sub)

    End Sub

    ''' <summary>
    ''' Every serial port this machine has, named but not described - this is the
    ''' fallback for when no probe announces itself.
    ''' </summary>
    Public Function AvailablePorts() As List(Of ProbePort)

        Dim ports As New List(Of ProbePort)

        Try
            For Each portName As String In SerialPort.GetPortNames()
                ports.Add(New ProbePort(portName, String.Empty))
            Next
        Catch
            ' No ports rather than an error: an empty list says the same thing.
        End Try

        Return ports

    End Function

    ' ========================================================================
    '  Finding a probe
    '
    '  A sweep of the USB device tree for serial ports whose product string
    '  starts with "BusMaster".
    '
    '  That string is iProduct out of the device's own descriptor, which Windows
    '  keeps as DEVPKEY_Device_BusReportedDeviceDesc. Nothing easier will do:
    '
    '    - the port's friendly name comes from the driver's INF, so an STM32
    '      running the inbox CDC driver is "USB Serial Device (COM7)" no matter
    '      what its firmware calls itself
    '    - the registry copy of the real name lives under a Properties key that
    '      cannot be read without elevation
    '
    '  SetupAPI needs neither, which is why this is interop rather than four
    '  lines of registry reading.
    ' ========================================================================

    ''' <summary>What a probe of ours calls itself over the bus.</summary>
    Public Const ProbeProductPrefix As String = "BusMaster"

    Private ReadOnly PortsClass As New Guid("4d36e978-e325-11ce-bfc1-08002be10318")

    ''' <summary>DEVPKEY_Device_BusReportedDeviceDesc - the name the device gives itself.</summary>
    Private ReadOnly BusReportedName As New DevicePropertyKey With {
        .Category = New Guid("540b947e-8b40-45bc-a8a2-6a0b894cbda2"),
        .Id = 4UI
    }

    Private Const DIGCF_PRESENT As UInteger = 2UI
    Private Const DICS_FLAG_GLOBAL As UInteger = 1UI
    Private Const DIREG_DEV As UInteger = 1UI
    Private Const KEY_READ As UInteger = &H20019UI
    Private Const ALL_PROFILES As UInteger = &HFFFFFFFFUI

    <StructLayout(LayoutKind.Sequential)>
    Private Structure DeviceInfo
        Public Size As UInteger
        Public ClassGuid As Guid
        Public Instance As UInteger
        Public Reserved As IntPtr
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure DevicePropertyKey
        Public Category As Guid
        Public Id As UInteger
    End Structure

    <DllImport("setupapi.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Private Function SetupDiGetClassDevs(ByRef classGuid As Guid, enumerator As String,
                                         parent As IntPtr, flags As UInteger) As IntPtr
    End Function

    <DllImport("setupapi.dll", SetLastError:=True)>
    Private Function SetupDiEnumDeviceInfo(deviceSet As IntPtr, index As UInteger,
                                           ByRef device As DeviceInfo) As Boolean
    End Function

    <DllImport("setupapi.dll", CharSet:=CharSet.Unicode, SetLastError:=True,
               EntryPoint:="SetupDiGetDevicePropertyW")>
    Private Function SetupDiGetDeviceProperty(deviceSet As IntPtr, ByRef device As DeviceInfo,
                                              ByRef propertyKey As DevicePropertyKey,
                                              ByRef propertyType As UInteger,
                                              buffer As Byte(), bufferSize As UInteger,
                                              ByRef required As UInteger, flags As UInteger) As Boolean
    End Function

    <DllImport("setupapi.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Private Function SetupDiOpenDevRegKey(deviceSet As IntPtr, ByRef device As DeviceInfo,
                                          scope As UInteger, profile As UInteger,
                                          keyType As UInteger, access As UInteger) As IntPtr
    End Function

    <DllImport("setupapi.dll")>
    Private Function SetupDiDestroyDeviceInfoList(deviceSet As IntPtr) As Boolean
    End Function

    <DllImport("advapi32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Private Function RegQueryValueEx(key As IntPtr, name As String, reserved As IntPtr,
                                     ByRef valueType As UInteger, data As StringBuilder,
                                     ByRef size As UInteger) As Integer
    End Function

    <DllImport("advapi32.dll")>
    Private Function RegCloseKey(key As IntPtr) As Integer
    End Function

    ''' <summary>
    ''' Every serial port on a plugged-in device that calls itself a BusMaster,
    ''' with the name it gave. Empty when there is none, which is not an error.
    ''' </summary>
    Public Function FindProbePorts() As List(Of ProbePort)

        Dim found As New List(Of ProbePort)
        Dim classGuid As Guid = PortsClass

        Dim devices As IntPtr = SetupDiGetClassDevs(classGuid, Nothing, IntPtr.Zero, DIGCF_PRESENT)
        If devices = New IntPtr(-1) Then Return found

        Try
            Dim device As New DeviceInfo
            device.Size = CUInt(Marshal.SizeOf(device))

            Dim index As UInteger = 0UI

            While SetupDiEnumDeviceInfo(devices, index, device)

                Dim product As String = ProductOf(devices, device)

                If product.StartsWith(ProbeProductPrefix, StringComparison.OrdinalIgnoreCase) Then

                    Dim portName As String = PortOf(devices, device)

                    If portName.Length > 0 AndAlso Not found.Any(Function(p) p.PortName = portName) Then
                        found.Add(New ProbePort(portName, product))
                    End If

                End If

                index += 1UI

            End While

        Finally
            SetupDiDestroyDeviceInfoList(devices)
        End Try

        Return found

    End Function

    ''' <summary>The device's own name for itself, or an empty string.</summary>
    Private Function ProductOf(devices As IntPtr, ByRef device As DeviceInfo) As String

        Dim propertyKey As DevicePropertyKey = BusReportedName
        Dim propertyType As UInteger = 0UI
        Dim required As UInteger = 0UI

        ' Asked once for the size, once for the value.
        SetupDiGetDeviceProperty(devices, device, propertyKey, propertyType, Nothing, 0UI, required, 0UI)

        If required = 0UI Then Return String.Empty

        Dim buffer(CInt(required) - 1) As Byte

        If Not SetupDiGetDeviceProperty(devices, device, propertyKey, propertyType,
                                        buffer, required, required, 0UI) Then Return String.Empty

        Return Encoding.Unicode.GetString(buffer).TrimEnd(ControlChars.NullChar)

    End Function

    ''' <summary>The COM name the device is bound to, or an empty string.</summary>
    Private Function PortOf(devices As IntPtr, ByRef device As DeviceInfo) As String

        Dim key As IntPtr = SetupDiOpenDevRegKey(devices, device, DICS_FLAG_GLOBAL,
                                                 ALL_PROFILES, DIREG_DEV, KEY_READ)

        If key = New IntPtr(-1) Then Return String.Empty

        Try
            Dim name As New StringBuilder(64)
            Dim size As UInteger = 128UI
            Dim valueType As UInteger = 0UI

            If RegQueryValueEx(key, "PortName", IntPtr.Zero, valueType, name, size) <> 0 Then
                Return String.Empty
            End If

            Return name.ToString()

        Finally
            RegCloseKey(key)
        End Try

    End Function

    ''' <summary>
    ''' Chooses which port the probe is on and remembers it. Anything already open
    ''' is closed first - the name is which port to use next, not which one is in
    ''' use now.
    ''' </summary>
    Public Sub ChoosePort(portName As String)

        Dim wanted As String = If(portName, String.Empty).Trim()

        If String.Equals(wanted, WantedPort, StringComparison.OrdinalIgnoreCase) Then Exit Sub

        Disconnect()

        AppSettings.Current.ProbePort = wanted
        AppSettings.Save()

        Complained = False

        If wanted.Length = 0 Then
            AppCore.SetStatus("Probe port cleared - the Terminal will turn lines straight round")
        Else
            AppCore.SetStatus("Probe port set to " & wanted)
        End If

    End Sub

    ''' <summary>
    ''' Opens the configured port if it is not open already. Reports a failure once
    ''' and then stays quiet, so a probe that is not plugged in does not fill the
    ''' transcript with the same complaint.
    ''' </summary>
    Public Function Connect() As Boolean

        If IsConnected Then Return True

        If WantedPort.Length = 0 Then Return False

        Try
            Port = New SerialPort(WantedPort, AppSettings.Current.ProbeBaud) With {
                .Encoding = Encoding.ASCII,
                .NewLine = LineEnding,
                .ReadTimeout = 500,
                .WriteTimeout = 500,
                .DtrEnable = True
            }

            AddHandler Port.DataReceived, AddressOf Port_DataReceived

            Port.Open()

            Complained = False
            Report("Probe connected on " & WantedPort)
            LinkChanged()

            Return True

        Catch ex As Exception

            Port = Nothing
            Complain("Probe port " & WantedPort & " would not open - " & ex.Message)
            LinkChanged()

            Return False

        End Try

    End Function

    ''' <summary>Closes the port. Safe to call when there is nothing open.</summary>
    Public Sub Disconnect()

        If Port Is Nothing Then Exit Sub

        Try
            RemoveHandler Port.DataReceived, AddressOf Port_DataReceived
            If Port.IsOpen Then Port.Close()
        Catch
            ' Closing a port that has already gone - a USB probe unplugged mid-run -
            ' throws, and there is nothing useful to do about it.
        End Try

        Port.Dispose()
        Port = Nothing

        Report("Probe disconnected")
        LinkChanged()

    End Sub

    Private Sub Transmit(text As String)

        ' No port named: turn the line straight round, so the Terminal is worth
        ' opening before any hardware exists. This is the whole of the stand-in.
        If WantedPort.Length = 0 Then
            OnUiThread(Sub() Absorb(text))
            Exit Sub
        End If

        If Not Connect() Then Exit Sub

        Try
            Port.Write(text)
        Catch ex As Exception
            Complain("Probe write failed - " & ex.Message)
            Disconnect()
        End Try

    End Sub

    ''' <summary>
    ''' The port's own event, raised on a worker thread. Nothing but the read
    ''' happens here: the text is handed to the UI thread before it is looked at,
    ''' which is what lets everything downstream ignore threading altogether.
    ''' </summary>
    Private Sub Port_DataReceived(sender As Object, e As SerialDataReceivedEventArgs)

        Dim arrived As String

        Try
            arrived = Port.ReadExisting()
        Catch
            Exit Sub
        End Try

        If arrived.Length = 0 Then Exit Sub

        OnUiThread(Sub() Absorb(arrived))

    End Sub

    ''' <summary>
    ''' Collects arriving text and lets whole lines through. A reply can be split
    ''' across any number of reads, so the tail is kept until its line finishes.
    ''' </summary>
    Private Sub Absorb(arrived As String)

        Arriving.Append(arrived)

        Dim pending As String = Arriving.ToString()
        Dim ends As Integer = pending.IndexOf(ControlChars.Lf)

        While ends >= 0

            ProbeReceive(pending.Substring(0, ends).TrimEnd(ControlChars.Cr))

            pending = pending.Substring(ends + 1)
            ends = pending.IndexOf(ControlChars.Lf)

        End While

        Arriving.Clear()
        Arriving.Append(pending)

    End Sub

    Private Sub OnUiThread(work As Action)

        Dim application As System.Windows.Application = System.Windows.Application.Current

        If application Is Nothing Then
            work()
            Exit Sub
        End If

        application.Dispatcher.BeginInvoke(DispatcherPriority.Normal, work)

    End Sub

    ''' <summary>
    ''' Says what went wrong, once. A probe that is not plugged in should not fill
    ''' the transcript with the same complaint. The next success clears the slate.
    ''' </summary>
    Private Sub Complain(message As String)

        If Complained Then Exit Sub

        Complained = True

        OnUiThread(Sub()
                       AppCore.SetStatus(message)
                       TerminalCore.ShowProblem(message)
                   End Sub)

    End Sub

    ''' <summary>Says what happened. Not a warning, so not in the warning colour.</summary>
    Private Sub Report(message As String)

        OnUiThread(Sub()
                       AppCore.SetStatus(message)
                       TerminalCore.ShowNotice(message)
                   End Sub)

    End Sub

End Module
