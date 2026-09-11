' ============================================================================
'  Modules\WindowSnap.vb
'
'  Sticks the tool windows - Event Log, Command, Terminal - to the edges of the
'  main window.
'
'  Drag one so its edge comes near an edge of the main window and let go: it
'  lands flush against that edge and takes the main window's measurement along
'  it, so a window dropped at the bottom ends up exactly as wide as the main
'  window. From then on it follows: move or resize the main window and every
'  window stuck to it comes along. Drag it away again and it is on its own.
'
'  Not docking. Nothing is reparented, the main window never gives up any of its
'  own space, and the tool windows stay ordinary top-level windows that can be
'  closed and moved as before.
'
'  Everything here is done in physical screen pixels through GetWindowRect and
'  SetWindowPos rather than in WPF's Left and Top:
'
'    - Left and Top do not follow a maximised window. They keep reporting where
'      it would go if it were restored, which would put the tool windows round a
'      rectangle that is not on screen.
'    - Two monitors at different scalings make device-independent arithmetic
'      subtly wrong the moment a window crosses between them. Physical pixels are
'      physical pixels on both.
'
'  It also means no feedback loop to guard against: SetWindowPos moves a window
'  without the drag messages this listens for, so a window being put in its place
'  never looks like a window being dragged.
'
'  One window to an edge. Sticking a second window to an edge that is taken
'  releases the first, which is left where it is rather than being moved - the
'  user asked for the new one to go there, not for the old one to go anywhere.
'
'  Two things are watched on a tool window, and the difference matters. The
'  window is placed on WM_EXITSIZEMOVE, when the mouse comes up, because snapping
'  a window that is still being held fights the hand holding it. The cue drawn on
'  the main window - see Modules\SnapHint.vb - is driven from WM_MOVING and
'  WM_SIZING instead, all the way through the drag, because its whole job is to
'  say what letting go would do before it is done.
' ============================================================================

Imports System.Runtime.InteropServices
Imports System.Windows.Interop
Imports System.Windows.Media

''' <summary>Which edge of the main window a tool window is stuck to.</summary>
Public Enum SnapEdge
    edge_None = 0
    edge_Left = 1
    edge_Right = 2
    edge_Top = 3
    edge_Bottom = 4
End Enum


Public Module WindowSnap

    ''' <summary>
    ''' How near an edge has to be, in pixels at 100%, before it takes. Scaled by
    ''' the window's own scaling, so it feels the same on a 200% display.
    ''' </summary>
    Private Const SnapDistance As Double = 20

    ''' <summary>
    ''' How much of the shared edge the two windows have to have in common before
    ''' one counts as being alongside the other. Without it, a window brought past
    ''' a corner would catch on an edge it is nowhere near.
    ''' </summary>
    Private Const MinimumOverlap As Double = 80


    Private Shell As Window

    Private ReadOnly Stuck As New List(Of SnapState)


    ''' <summary>One tool window and the edge it is stuck to, if any.</summary>
    Private NotInheritable Class SnapState

        Public ReadOnly Window As Window
        Public Edge As SnapEdge = SnapEdge.edge_None

        Public Sub New(tool As Window)
            Window = tool
        End Sub

    End Class


    ' ========================================================================
    '  Setting up
    ' ========================================================================

    ''' <summary>
    ''' The window everything else sticks to. Watched for every kind of move there
    ''' is - dragged, resized, maximised, restored - so that whatever happens to it
    ''' happens to the windows stuck to it.
    ''' </summary>
    Public Sub Attach(main As Window)

        Shell = main

        Hook(main, AddressOf ShellMessage)

    End Sub

    ''' <summary>
    ''' Lets a tool window stick to the main one. Called when the window is built,
    ''' which is before it is ever shown - the hook waits for the handle.
    ''' </summary>
    Public Sub Follow(tool As Window)

        If tool Is Nothing Then Exit Sub

        Stuck.Add(New SnapState(tool))

        Hook(tool, AddressOf ToolMessage)

    End Sub

    ''' <summary>
    ''' Puts a message hook on a window, waiting for its handle if it has not been
    ''' shown yet. A WPF window has no handle until SourceInitialized.
    ''' </summary>
    Private Sub Hook(window As Window, handler As HwndSourceHook)

        Dim source As HwndSource = TryCast(PresentationSource.FromVisual(window), HwndSource)

        If source IsNot Nothing Then
            source.AddHook(handler)
            Exit Sub
        End If

        AddHandler window.SourceInitialized,
            Sub()
                Dim made As HwndSource = TryCast(PresentationSource.FromVisual(window), HwndSource)
                If made IsNot Nothing Then made.AddHook(handler)
            End Sub

    End Sub


    ' ========================================================================
    '  Listening
    ' ========================================================================

    ''' <summary>
    ''' The main window has been moved, resized, maximised or restored. Everything
    ''' stuck to it is put back where it belongs.
    ''' </summary>
    Private Function ShellMessage(handle As IntPtr, message As Integer, wparam As IntPtr,
                                  lparam As IntPtr, ByRef handled As Boolean) As IntPtr

        If message = WM_WINDOWPOSCHANGED Then ReflowAll()

        Return IntPtr.Zero

    End Function

    ''' <summary>
    ''' A tool window is being dragged or resized, or has just been let go of.
    '''
    ''' Nothing is moved until the mouse comes up: snapping while it is still down
    ''' fights the hand holding it, and a window that jumps out from under the
    ''' pointer is worse than one that does not snap at all. What does happen
    ''' during the drag is the cue on the main window, which has to be live or it
    ''' is telling the user nothing they could act on.
    ''' </summary>
    Private Function ToolMessage(handle As IntPtr, message As Integer, wparam As IntPtr,
                                 lparam As IntPtr, ByRef handled As Boolean) As IntPtr

        Select Case message

            Case WM_MOVING, WM_SIZING
                Preview(handle, lparam)

            Case WM_EXITSIZEMOVE
                SnapHint.Hide()
                Settle(handle)

        End Select

        Return IntPtr.Zero

    End Function


    ' ========================================================================
    '  Deciding
    ' ========================================================================

    ''' <summary>
    ''' Says on the main window what letting go here would do. Called on every
    ''' mouse move of a drag, and on every step of a resize, so it does as little
    ''' as it can: read two rectangles, ask the same question Settle will ask, and
    ''' hand the answer to SnapHint, which draws nothing new if nothing changed.
    ''' </summary>
    Private Sub Preview(handle As IntPtr, lparam As IntPtr)

        If Shell Is Nothing Then Exit Sub
        If lparam = IntPtr.Zero Then Exit Sub

        Dim state As SnapState = StateOf(handle)
        If state Is Nothing Then Exit Sub

        ' Where the window is about to be, not where it is. Both messages arrive
        ' before the move with the rectangle Windows is proposing, so reading the
        ' window itself here would be a frame behind the pointer.
        Dim toolRect As RECT = Marshal.PtrToStructure(Of RECT)(lparam)

        Dim shellRect As RECT
        If Not GetWindowRect(HandleOf(Shell), shellRect) Then Exit Sub

        Dim edge As SnapEdge = NearestEdge(toolRect, shellRect,
                                           VisualTreeHelper.GetDpi(state.Window).DpiScaleX)

        SnapHint.Show(Shell, edge)

    End Sub

    ''' <summary>
    ''' Works out where the tool window that was just let go of belongs now, and
    ''' puts it there.
    ''' </summary>
    Private Sub Settle(handle As IntPtr)

        If Shell Is Nothing Then Exit Sub

        Dim state As SnapState = StateOf(handle)
        If state Is Nothing Then Exit Sub

        Dim shellRect As RECT
        Dim toolRect As RECT

        If Not GetWindowRect(HandleOf(Shell), shellRect) Then Exit Sub
        If Not GetWindowRect(handle, toolRect) Then Exit Sub

        ' A window being resized while stuck keeps the edge it is stuck to: the
        ' free edge is the one that moved, and Reflow puts the other three back.
        Dim edge As SnapEdge = NearestEdge(toolRect, shellRect, VisualTreeHelper.GetDpi(state.Window).DpiScaleX)

        If edge = SnapEdge.edge_None Then
            state.Edge = SnapEdge.edge_None
            Exit Sub
        End If

        ' An edge holds one window.
        For Each other As SnapState In Stuck
            If other IsNot state AndAlso other.Edge = edge Then other.Edge = SnapEdge.edge_None
        Next

        state.Edge = edge
        Reflow(state, shellRect)

    End Sub

    ''' <summary>
    ''' The edge of the main window this one has been brought near enough to, or
    ''' edge_None. Where two would do, the nearer one wins.
    ''' </summary>
    Private Function NearestEdge(tool As RECT, shell As RECT, scaling As Double) As SnapEdge

        Dim reach As Double = SnapDistance * Math.Max(scaling, 1)
        Dim overlap As Double = MinimumOverlap * Math.Max(scaling, 1)

        Dim best As SnapEdge = SnapEdge.edge_None
        Dim nearest As Double = Double.MaxValue

        Dim sideBySide As Boolean = Overlaps(tool.Top, tool.Bottom, shell.Top, shell.Bottom, overlap)
        Dim stacked As Boolean = Overlaps(tool.Left, tool.Right, shell.Left, shell.Right, overlap)

        If stacked Then

            Dim below As Double = Math.Abs(tool.Top - shell.Bottom)
            If below <= reach AndAlso below < nearest Then
                nearest = below
                best = SnapEdge.edge_Bottom
            End If

            Dim above As Double = Math.Abs(tool.Bottom - shell.Top)
            If above <= reach AndAlso above < nearest Then
                nearest = above
                best = SnapEdge.edge_Top
            End If

        End If

        If sideBySide Then

            Dim rightOf As Double = Math.Abs(tool.Left - shell.Right)
            If rightOf <= reach AndAlso rightOf < nearest Then
                nearest = rightOf
                best = SnapEdge.edge_Right
            End If

            Dim leftOf As Double = Math.Abs(tool.Right - shell.Left)
            If leftOf <= reach AndAlso leftOf < nearest Then
                nearest = leftOf
                best = SnapEdge.edge_Left
            End If

        End If

        Return best

    End Function

    ''' <summary>How much two spans share, against how much they have to.</summary>
    Private Function Overlaps(fromA As Integer, toA As Integer,
                              fromB As Integer, toB As Integer, wanted As Double) As Boolean

        ' Not "shared" - that is a VB keyword.
        Dim common As Integer = Math.Min(toA, toB) - Math.Max(fromA, fromB)

        ' A window smaller than the requirement only has to overlap all of itself.
        Dim needed As Double = Math.Min(wanted, Math.Min(toA - fromA, toB - fromB))

        Return common >= needed

    End Function


    ' ========================================================================
    '  Placing
    ' ========================================================================

    Private Sub ReflowAll()

        If Shell Is Nothing Then Exit Sub

        Dim shellRect As RECT
        If Not GetWindowRect(HandleOf(Shell), shellRect) Then Exit Sub

        For Each state As SnapState In Stuck
            Reflow(state, shellRect)
        Next

    End Sub

    ''' <summary>
    ''' Puts one tool window flush against its edge, matching the main window along
    ''' that edge and keeping its own measurement across it - a window stuck to the
    ''' bottom takes the main window's width and keeps its own height.
    ''' </summary>
    Private Sub Reflow(state As SnapState, shell As RECT)

        If state.Edge = SnapEdge.edge_None Then Exit Sub
        If Not state.Window.IsVisible Then Exit Sub

        Dim handle As IntPtr = HandleOf(state.Window)
        If handle = IntPtr.Zero Then Exit Sub

        Dim tool As RECT
        If Not GetWindowRect(handle, tool) Then Exit Sub

        Dim width As Integer = tool.Right - tool.Left
        Dim height As Integer = tool.Bottom - tool.Top

        Dim x As Integer
        Dim y As Integer

        Select Case state.Edge

            Case SnapEdge.edge_Bottom
                x = shell.Left
                y = shell.Bottom
                width = shell.Right - shell.Left

            Case SnapEdge.edge_Top
                x = shell.Left
                y = shell.Top - height
                width = shell.Right - shell.Left

            Case SnapEdge.edge_Right
                x = shell.Right
                y = shell.Top
                height = shell.Bottom - shell.Top

            Case SnapEdge.edge_Left
                x = shell.Left - width
                y = shell.Top
                height = shell.Bottom - shell.Top

        End Select

        SetWindowPos(handle, IntPtr.Zero, x, y, width, height,
                     SWP_NOZORDER Or SWP_NOACTIVATE)

    End Sub


    ' ========================================================================
    '  Small helpers
    ' ========================================================================

    Private Function StateOf(handle As IntPtr) As SnapState

        For Each state As SnapState In Stuck
            If HandleOf(state.Window) = handle Then Return state
        Next

        Return Nothing

    End Function

    Private Function HandleOf(window As Window) As IntPtr

        If window Is Nothing Then Return IntPtr.Zero

        Dim helper As New WindowInteropHelper(window)

        Return helper.Handle

    End Function


    ' ========================================================================
    '  Win32
    ' ========================================================================

    Private Const WM_SIZING As Integer = &H214
    Private Const WM_MOVING As Integer = &H216
    Private Const WM_EXITSIZEMOVE As Integer = &H232
    Private Const WM_WINDOWPOSCHANGED As Integer = &H47

    Private Const SWP_NOZORDER As UInteger = &H4
    Private Const SWP_NOACTIVATE As UInteger = &H10

    <StructLayout(LayoutKind.Sequential)>
    Private Structure RECT
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure

    <DllImport("user32.dll", SetLastError:=True)>
    Private Function GetWindowRect(handle As IntPtr, ByRef area As RECT) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Function SetWindowPos(handle As IntPtr, after As IntPtr,
                                  x As Integer, y As Integer,
                                  width As Integer, height As Integer,
                                  flags As UInteger) As Boolean
    End Function

End Module
