' ============================================================================
'  Modules\EventLogModel.vb
'
'  One line of the event log, and the marker that decides how it is coloured.
'
'  The marker in the second column is the whole of the line's formatting: nothing
'  else is stored, so a log read back from CSV comes out looking exactly as it did
'  when it was written.
'
'      L   label      yellow on black, and it pushes a fresh line in below it
'      S   start      bright green
'      P   stop       bright red
'      D   data       bright blue
'      /   comment    dark green, over whatever the row's stripe is
'      (blank)        white
' ============================================================================

Imports System.ComponentModel

Public Module EventLogMarkers

    Public Const Label As String = "L"
    Public Const StartCondition As String = "S"
    Public Const StopCondition As String = "P"
    Public Const DataByte As String = "D"
    Public Const Comment As String = "/"

    ''' <summary>Everything the second column will accept, besides a blank.</summary>
    Public Const Allowed As String = "LSPD/"

    ''' <summary>Resource key of the ink a marker paints the rest of its row in.</summary>
    Public Function ForegroundKey(marker As String) As String

        Select Case If(marker, String.Empty).Trim().ToUpperInvariant()
            Case Label
                Return "Brush_Log_Label"
            Case StartCondition
                Return "Brush_Log_Start"
            Case StopCondition
                Return "Brush_Log_Stop"
            Case DataByte
                Return "Brush_Log_Data"
            Case Comment
                Return "Brush_Log_Comment"
            Case Else
                Return "Brush_Log_Plain"
        End Select

    End Function

End Module


''' <summary>
''' One row of the grid. Number is kept in step by the window; Marker drives the
''' colours, which is why setting it announces the brush properties as well.
''' </summary>
Public Class EventLogRow
    Implements INotifyPropertyChanged

    Public Event PropertyChanged As PropertyChangedEventHandler _
        Implements INotifyPropertyChanged.PropertyChanged

    Private m_Number As Integer = 0
    Private m_Marker As String = String.Empty
    ' Host, Reg and Value all hold a number from 0 to 255, kept as text so a
    ' half-typed cell survives.
    Private m_Host As String = String.Empty
    Private m_Reg As String = String.Empty
    Private m_Value As String = String.Empty
    Private m_Comment As String = String.Empty
    Private m_IsBreakpoint As Boolean = False

    ''' <summary>Sequential line number. Renumbered whenever rows come or go.</summary>
    Public Property Number As Integer
        Get
            Return m_Number
        End Get
        Set(value As Integer)
            If m_Number = value Then Exit Property
            m_Number = value
            Announce(NameOf(Number))
        End Set
    End Property

    Public Property Marker As String
        Get
            Return m_Marker
        End Get
        Set(value As String)
            Dim tidied As String = If(value, String.Empty).Trim().ToUpperInvariant()
            If tidied.Length > 1 Then tidied = tidied.Substring(0, 1)

            If String.Equals(m_Marker, tidied, StringComparison.Ordinal) Then Exit Property

            m_Marker = tidied

            Announce(NameOf(Marker))
            ' The colours are derived, so they change with it.
            Announce(NameOf(IsLabel))
            Announce(NameOf(Foreground))
            Announce(NameOf(CommentForeground))
        End Set
    End Property

    Public Property Host As String
        Get
            Return m_Host
        End Get
        Set(value As String)
            If String.Equals(m_Host, value, StringComparison.Ordinal) Then Exit Property
            m_Host = If(value, String.Empty)
            Announce(NameOf(Host))
        End Set
    End Property

    Public Property Reg As String
        Get
            Return m_Reg
        End Get
        Set(value As String)
            If String.Equals(m_Reg, value, StringComparison.Ordinal) Then Exit Property
            m_Reg = If(value, String.Empty)
            Announce(NameOf(Reg))
        End Set
    End Property

    Public Property Value As String
        Get
            Return m_Value
        End Get
        Set(newValue As String)
            If String.Equals(m_Value, newValue, StringComparison.Ordinal) Then Exit Property
            m_Value = If(newValue, String.Empty)
            Announce(NameOf(Value))
        End Set
    End Property

    Public Property Comment As String
        Get
            Return m_Comment
        End Get
        Set(value As String)
            If String.Equals(m_Comment, value, StringComparison.Ordinal) Then Exit Property
            m_Comment = If(value, String.Empty)
            Announce(NameOf(Comment))
        End Set
    End Property

    ''' <summary>
    ''' A breakpoint, set by double-clicking the line number and shown by colouring
    ''' that cell red. Deliberately not part of the file: a breakpoint belongs to
    ''' the session, not to the recording.
    ''' </summary>
    Public Property IsBreakpoint As Boolean
        Get
            Return m_IsBreakpoint
        End Get
        Set(value As Boolean)
            If m_IsBreakpoint = value Then Exit Property
            m_IsBreakpoint = value
            Announce(NameOf(IsBreakpoint))
        End Set
    End Property

    ''' <summary>A label line, which is the one marker that also changes the backdrop.</summary>
    Public ReadOnly Property IsLabel As Boolean
        Get
            Return m_Marker = EventLogMarkers.Label
        End Get
    End Property

    ''' <summary>
    ''' This row's ink, resolved from the theme. Handed straight to the cell styles,
    ''' which saves every one of them needing a converter.
    ''' </summary>
    Public ReadOnly Property Foreground As Brush
        Get
            Return Ink(EventLogMarkers.ForegroundKey(m_Marker))
        End Get
    End Property

    ''' <summary>
    ''' The comment column's ink. Comments read as green unless the line carries a
    ''' marker, in which case the marker wins and the whole line is one colour -
    ''' including "/", whose darker green is the point of choosing it.
    ''' </summary>
    Public ReadOnly Property CommentForeground As Brush
        Get
            If m_Marker.Length > 0 Then Return Foreground

            Return Ink("Brush_Log_Comment_Column")
        End Get
    End Property

    ''' <summary>A theme brush by key, or white if the theme has not loaded.</summary>
    Private Shared Function Ink(key As String) As Brush

        If System.Windows.Application.Current IsNot Nothing Then
            Dim found As Object = System.Windows.Application.Current.TryFindResource(key)
            If TypeOf found Is Brush Then Return DirectCast(found, Brush)
        End If

        Return Brushes.White

    End Function

    ''' <summary>True when nothing has been typed anywhere on the line.</summary>
    Public Function IsBlank() As Boolean

        Return m_Marker.Length = 0 AndAlso
               String.IsNullOrWhiteSpace(m_Host) AndAlso
               String.IsNullOrWhiteSpace(m_Reg) AndAlso
               String.IsNullOrWhiteSpace(m_Value) AndAlso
               String.IsNullOrWhiteSpace(m_Comment)

    End Function

    Private Sub Announce(propertyName As String)

        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))

    End Sub

End Class

