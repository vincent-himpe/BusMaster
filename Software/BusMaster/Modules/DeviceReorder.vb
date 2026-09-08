' ============================================================================
'  Modules\DeviceReorder.vb
'
'  Dragging a device panel to a new place in the stack.
'
'  DevicePanel says when the user has taken hold of its grab bars; everything
'  after that happens here, because the stack belongs to the workspace and not to
'  any one panel.
'
'  The drag runs on a mouse capture rather than DragDrop.DoDragDrop. Nothing ever
'  leaves this one StackPanel, so there is no drag data to carry and no need for a
'  nested message loop - and a capture leaves the auto-scroll and the Escape key
'  simple to do.
'
'  Nothing moves until the drop. The blue line is an adorner floating over the
'  stack, so the panels stay where they are and the arithmetic that works out
'  where the pointer is does not shift under itself.
' ============================================================================

Imports System.Globalization
Imports System.Windows.Documents
Imports System.Windows.Threading

Public Module DeviceReorder

    ''' <summary>Gap between panels, from the margin NewDevicePanel gives them.</summary>
    Private Const PanelGap As Double = 6

    ''' <summary>How close to an edge of the workspace starts an auto-scroll.</summary>
    Private Const EdgeZone As Double = 28

    ''' <summary>Pixels per tick while auto-scrolling, and how often a tick comes.</summary>
    Private Const ScrollStep As Double = 14
    Private Const ScrollInterval As Integer = 40

    ''' <summary>The panel being dragged is dimmed, so it is clear what is moving.</summary>
    Private Const DraggingOpacity As Double = 0.55

    Private Panel As DevicePanel
    Private Stack As StackPanel
    Private Scroller As ScrollViewer
    Private Shell As MainWindow

    Private Line As InsertionAdorner
    Private Layer As AdornerLayer
    Private Scrolling As DispatcherTimer

    ''' <summary>Where the panel started, and where it would land - both gap indices.</summary>
    Private FromIndex As Integer = -1
    Private ToIndex As Integer = -1


    ' ========================================================================
    '  Starting and stopping
    ' ========================================================================

    ''' <summary>
    ''' Takes over a drag the panel has just started. Does nothing if there is only
    ''' one device on screen, or if a drag is somehow already running.
    ''' </summary>
    Public Sub Begin(source As Object)

        Dim moving As DevicePanel = TryCast(source, DevicePanel)
        If moving Is Nothing OrElse Panel IsNot Nothing Then Exit Sub

        Shell = AppCore.MainShell
        If Shell Is Nothing Then Exit Sub

        Stack = Shell.stk_Devices
        Scroller = Shell.scr_Workspace

        If Stack.Children.Count < 2 Then Exit Sub

        FromIndex = Stack.Children.IndexOf(moving)
        If FromIndex < 0 Then Exit Sub

        Layer = AdornerLayer.GetAdornerLayer(Stack)
        If Layer Is Nothing Then Exit Sub

        Panel = moving
        ToIndex = FromIndex

        Line = New InsertionAdorner(Stack)
        Layer.Add(Line)

        Panel.Opacity = DraggingOpacity

        AddHandler Stack.PreviewMouseMove, AddressOf Stack_PreviewMouseMove
        AddHandler Stack.PreviewMouseLeftButtonUp, AddressOf Stack_PreviewMouseLeftButtonUp
        AddHandler Stack.LostMouseCapture, AddressOf Stack_LostMouseCapture
        AddHandler Shell.PreviewKeyDown, AddressOf Shell_PreviewKeyDown

        Stack.CaptureMouse()

        Scrolling = New DispatcherTimer With {
            .Interval = TimeSpan.FromMilliseconds(ScrollInterval)
        }
        AddHandler Scrolling.Tick, AddressOf Scrolling_Tick
        Scrolling.Start()

        Track()

        AppCore.SetStatus("Drop " & Panel.DeviceName & " on the blue line, or press Escape")

    End Sub

    ''' <summary>
    ''' Puts everything back the way it was found. Safe to call twice, which matters
    ''' because releasing the capture raises LostMouseCapture in the middle of it.
    ''' </summary>
    Private Sub Finish()

        If Panel Is Nothing Then Exit Sub

        Dim moved As DevicePanel = Panel

        ' Cleared first: LostMouseCapture comes back through here.
        Panel = Nothing

        If Scrolling IsNot Nothing Then
            Scrolling.Stop()
            RemoveHandler Scrolling.Tick, AddressOf Scrolling_Tick
            Scrolling = Nothing
        End If

        RemoveHandler Stack.PreviewMouseMove, AddressOf Stack_PreviewMouseMove
        RemoveHandler Stack.PreviewMouseLeftButtonUp, AddressOf Stack_PreviewMouseLeftButtonUp
        RemoveHandler Stack.LostMouseCapture, AddressOf Stack_LostMouseCapture
        RemoveHandler Shell.PreviewKeyDown, AddressOf Shell_PreviewKeyDown

        If Stack.IsMouseCaptured Then Stack.ReleaseMouseCapture()

        If Line IsNot Nothing AndAlso Layer IsNot Nothing Then Layer.Remove(Line)

        Line = Nothing
        Layer = Nothing

        moved.Opacity = 1

    End Sub


    ' ========================================================================
    '  Following the pointer
    ' ========================================================================

    Private Sub Stack_PreviewMouseMove(sender As Object, e As MouseEventArgs)

        If Panel Is Nothing Then Exit Sub

        Track()
        e.Handled = True

    End Sub

    ''' <summary>Works out which gap the pointer is nearest and moves the line to it.</summary>
    Private Sub Track()

        If Panel Is Nothing OrElse Line Is Nothing Then Exit Sub

        ToIndex = NearestGap(Mouse.GetPosition(Stack).Y)
        Line.LineY = GapPosition(ToIndex)

    End Sub

    ''' <summary>
    ''' The gap the pointer is closest to, as an insertion index: 0 is above the
    ''' first panel, Count is below the last.
    '''
    ''' Anywhere over a panel counts - its top half means the gap above it, its
    ''' bottom half the gap below. Asking the user to hit the six pixels actually
    ''' between two panels would be needlessly hard.
    ''' </summary>
    Private Function NearestGap(y As Double) As Integer

        For index As Integer = 0 To Stack.Children.Count - 1

            Dim child As FrameworkElement = TryCast(Stack.Children(index), FrameworkElement)
            If child Is Nothing Then Continue For

            Dim top As Double = child.TranslatePoint(New Point(0, 0), Stack).Y
            Dim height As Double = child.ActualHeight

            If y < top + (height / 2) Then Return index
            If y < top + height + PanelGap Then Return index + 1

        Next

        Return Stack.Children.Count

    End Function

    ''' <summary>Where the line sits for an insertion index, in the stack's own space.</summary>
    Private Function GapPosition(index As Integer) As Double

        If Stack.Children.Count = 0 Then Return 0

        Dim y As Double

        If index >= Stack.Children.Count Then

            Dim last As FrameworkElement = TryCast(Stack.Children(Stack.Children.Count - 1), FrameworkElement)
            If last Is Nothing Then Return 0

            y = last.TranslatePoint(New Point(0, 0), Stack).Y + last.ActualHeight + (PanelGap / 2)

        Else

            Dim child As FrameworkElement = TryCast(Stack.Children(index), FrameworkElement)
            If child Is Nothing Then Return 0

            y = child.TranslatePoint(New Point(0, 0), Stack).Y - (PanelGap / 2)

        End If

        ' Half a line width in from either end, so a drop at the very top or bottom
        ' still draws a whole line.
        Dim inset As Double = InsertionAdorner.LineThickness / 2

        Return Math.Max(inset, Math.Min(Stack.ActualHeight - inset, y))

    End Function

    ''' <summary>
    ''' Scrolls the workspace while the pointer is held near its top or bottom edge,
    ''' so a device can be dragged past the end of what is on screen.
    ''' </summary>
    Private Sub Scrolling_Tick(sender As Object, e As EventArgs)

        If Panel Is Nothing Then Exit Sub

        Dim where As Point = Mouse.GetPosition(Scroller)

        If where.Y < EdgeZone Then
            Scroller.ScrollToVerticalOffset(Scroller.VerticalOffset - ScrollStep)
        ElseIf where.Y > Scroller.ActualHeight - EdgeZone Then
            Scroller.ScrollToVerticalOffset(Scroller.VerticalOffset + ScrollStep)
        Else
            Exit Sub
        End If

        ' The pointer has not moved but the panels under it have.
        Track()

    End Sub


    ' ========================================================================
    '  Finishing
    ' ========================================================================

    Private Sub Stack_PreviewMouseLeftButtonUp(sender As Object, e As MouseButtonEventArgs)

        Dim moving As DevicePanel = Panel
        Dim landing As Integer = ToIndex
        Dim started As Integer = FromIndex

        Finish()

        e.Handled = True

        If moving Is Nothing Then Exit Sub

        ' The gap just above and the gap just below are both where it already is.
        If landing = started OrElse landing = started + 1 Then
            AppCore.SetStatus("Left " & moving.DeviceName & " where it was")
            Exit Sub
        End If

        ' Taking it out shifts everything below up one, so a landing further down
        ' the stack has to come back by the same one.
        Dim insertAt As Integer = landing
        If landing > started Then insertAt -= 1

        Stack.Children.Remove(moving)
        Stack.Children.Insert(insertAt, moving)

        ' The order on screen is the order that gets saved, so this is a change to
        ' the project like any other.
        AppCore.MarkModified()

        AppCore.SetStatus("Moved " & moving.DeviceName & " to position " &
                          (insertAt + 1).ToString(CultureInfo.InvariantCulture))

    End Sub

    Private Sub Shell_PreviewKeyDown(sender As Object, e As KeyEventArgs)

        If e.Key <> Key.Escape Then Exit Sub

        Finish()
        e.Handled = True

        AppCore.SetStatus("Move cancelled")

    End Sub

    ''' <summary>
    ''' Anything that takes the capture away - another window coming forward, a
    ''' message box - abandons the move rather than dropping it somewhere the user
    ''' cannot see.
    ''' </summary>
    Private Sub Stack_LostMouseCapture(sender As Object, e As MouseEventArgs)

        Finish()

    End Sub

End Module


''' <summary>
''' The line that says where a dragged device would land: drawn over the stack, so
''' nothing has to move until the drop.
''' </summary>
Public Class InsertionAdorner
    Inherits Adorner

    Public Const LineThickness As Double = 4

    Private ReadOnly Ink As Pen
    Private m_LineY As Double = 0

    Public Sub New(adorned As UIElement)

        MyBase.New(adorned)

        ' The pointer has to reach the panels underneath.
        IsHitTestVisible = False

        Ink = New Pen(LineBrush(), LineThickness)
        Ink.Freeze()

    End Sub

    ''' <summary>Where the line is, in the adorned element's own coordinates.</summary>
    Public Property LineY As Double
        Get
            Return m_LineY
        End Get
        Set(value As Double)
            If m_LineY = value Then Exit Property
            m_LineY = value
            InvalidateVisual()
        End Set
    End Property

    Protected Overrides Sub OnRender(context As DrawingContext)

        MyBase.OnRender(context)

        Dim host As FrameworkElement = TryCast(AdornedElement, FrameworkElement)
        If host Is Nothing Then Exit Sub

        context.DrawLine(Ink, New Point(0, m_LineY), New Point(host.ActualWidth, m_LineY))

    End Sub

    ''' <summary>
    ''' The same blue the function name is written in, so the line reads as
    ''' belonging to the panel being moved.
    ''' </summary>
    Private Shared Function LineBrush() As Brush

        If System.Windows.Application.Current IsNot Nothing Then
            Dim found As Object = System.Windows.Application.Current.TryFindResource("Brush_Function_Name")
            If TypeOf found Is Brush Then Return DirectCast(found, Brush)
        End If

        Return Brushes.DeepSkyBlue

    End Function

End Class
