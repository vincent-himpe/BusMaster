' ============================================================================
'  Modules\DeviceFileModel.vb
'
'  What a .DEV device file contains, and the editable row behind it.
'
'  Same house rules as the settings and project files: strings, numbers and
'  Booleans only, written indented so the file reads and diffs by hand.
' ============================================================================

Imports System.ComponentModel

''' <summary>Which of the three jobs the Device Editor was opened to do.</summary>
Public Enum DeviceEditorMode
    mode_New = 0
    mode_Edit = 1
    mode_Clone = 2
End Enum


''' <summary>The shape of a .DEV file.</summary>
Public Class DeviceFileData
    Public Property SchemaVersion As Integer = 1
    Public Property DeviceName As String = String.Empty
    Public Property DeviceDescription As String = String.Empty
    Public Property BaseAddress As String = String.Empty
    Public Property AddressBits As String = String.Empty
    Public Property Registers As New List(Of DeviceRegisterData)
End Class


''' <summary>
''' One register in a .DEV file. D7 to D0 are the bit names, written out one key
''' each so the file says which bit is which without anyone having to count.
'''
''' RegisterGroup ties consecutive addresses together into one wider register: two
''' rows sharing a group are a 16 bit register, four are 32, and a blank group is
''' the ordinary 8 bit case. It was added after the fact, so a file written without
''' it still loads - every row simply comes back ungrouped.
''' </summary>
Public Class DeviceRegisterData
    Public Property RegisterAddress As Integer = 0
    Public Property RegisterGroup As String = String.Empty
    Public Property RegisterName As String = String.Empty
    Public Property D7 As String = String.Empty
    Public Property D6 As String = String.Empty
    Public Property D5 As String = String.Empty
    Public Property D4 As String = String.Empty
    Public Property D3 As String = String.Empty
    Public Property D2 As String = String.Empty
    Public Property D1 As String = String.Empty
    Public Property D0 As String = String.Empty
End Class


''' <summary>
''' One row of the editor's grid. Everything is text, because a half typed row has
''' to survive until Save gets a chance to look at it.
'''
''' Raises PropertyChanged only when a value genuinely changes, which is what lets
''' the editor tell a real edit from the user tabbing through a cell and leaving it
''' alone.
''' </summary>
Public Class DeviceRegisterRow
    Implements INotifyPropertyChanged

    Public Event PropertyChanged As PropertyChangedEventHandler _
        Implements INotifyPropertyChanged.PropertyChanged

    Private m_RegisterAddress As String = String.Empty
    Private m_RegisterGroup As String = String.Empty
    Private m_RegisterName As String = String.Empty
    Private m_ShadeIndex As Integer = 0
    Private m_D7 As String = String.Empty
    Private m_D6 As String = String.Empty
    Private m_D5 As String = String.Empty
    Private m_D4 As String = String.Empty
    Private m_D3 As String = String.Empty
    Private m_D2 As String = String.Empty
    Private m_D1 As String = String.Empty
    Private m_D0 As String = String.Empty

    Public Property RegisterAddress As String
        Get
            Return m_RegisterAddress
        End Get
        Set(value As String)
            Assign(m_RegisterAddress, value, NameOf(RegisterAddress))
        End Set
    End Property

    ''' <summary>
    ''' Group tag. Blank means an ordinary 8 bit register; a tag shared with the row
    ''' at the next address makes the pair one wider register.
    ''' </summary>
    Public Property RegisterGroup As String
        Get
            Return m_RegisterGroup
        End Get
        Set(value As String)
            Assign(m_RegisterGroup, value, NameOf(RegisterGroup))
        End Set
    End Property

    Public Property RegisterName As String
        Get
            Return m_RegisterName
        End Get
        Set(value As String)
            Assign(m_RegisterName, value, NameOf(RegisterName))
        End Set
    End Property

    ''' <summary>
    ''' Which of the two row backgrounds this row draws itself in - 0 black, 1 at
    ''' 95% black. Worked out by the editor, not typed in and never saved: it is
    ''' here only because a cell style can bind to it. Rows of the same group get
    ''' the same number, so a group reads as one block.
    ''' </summary>
    Public Property ShadeIndex As Integer
        Get
            Return m_ShadeIndex
        End Get
        Set(value As Integer)
            If m_ShadeIndex = value Then Exit Property

            m_ShadeIndex = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(ShadeIndex)))
        End Set
    End Property

    Public Property D7 As String
        Get
            Return m_D7
        End Get
        Set(value As String)
            Assign(m_D7, value, NameOf(D7))
        End Set
    End Property

    Public Property D6 As String
        Get
            Return m_D6
        End Get
        Set(value As String)
            Assign(m_D6, value, NameOf(D6))
        End Set
    End Property

    Public Property D5 As String
        Get
            Return m_D5
        End Get
        Set(value As String)
            Assign(m_D5, value, NameOf(D5))
        End Set
    End Property

    Public Property D4 As String
        Get
            Return m_D4
        End Get
        Set(value As String)
            Assign(m_D4, value, NameOf(D4))
        End Set
    End Property

    Public Property D3 As String
        Get
            Return m_D3
        End Get
        Set(value As String)
            Assign(m_D3, value, NameOf(D3))
        End Set
    End Property

    Public Property D2 As String
        Get
            Return m_D2
        End Get
        Set(value As String)
            Assign(m_D2, value, NameOf(D2))
        End Set
    End Property

    Public Property D1 As String
        Get
            Return m_D1
        End Get
        Set(value As String)
            Assign(m_D1, value, NameOf(D1))
        End Set
    End Property

    Public Property D0 As String
        Get
            Return m_D0
        End Get
        Set(value As String)
            Assign(m_D0, value, NameOf(D0))
        End Set
    End Property

    ''' <summary>
    ''' The eight bit cells by column position: 0 is D7, 7 is D0. Lets the editor
    ''' work on a bit column without caring which property sits behind it.
    ''' </summary>
    Public Property BitName(position As Integer) As String
        Get
            Select Case position
                Case 0
                    Return D7
                Case 1
                    Return D6
                Case 2
                    Return D5
                Case 3
                    Return D4
                Case 4
                    Return D3
                Case 5
                    Return D2
                Case 6
                    Return D1
                Case 7
                    Return D0
                Case Else
                    Return String.Empty
            End Select
        End Get
        Set(value As String)
            Select Case position
                Case 0
                    D7 = value
                Case 1
                    D6 = value
                Case 2
                    D5 = value
                Case 3
                    D4 = value
                Case 4
                    D3 = value
                Case 5
                    D2 = value
                Case 6
                    D1 = value
                Case 7
                    D0 = value
            End Select
        End Set
    End Property

    ''' <summary>Stores a value and announces it, but only if it really changed.</summary>
    Private Sub Assign(ByRef field As String, value As String, propertyName As String)

        If String.Equals(field, value, StringComparison.Ordinal) Then Exit Sub

        field = value
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))

    End Sub

    ''' <summary>True when every cell is empty. Blank rows are never saved.</summary>
    Public Function IsBlank() As Boolean

        Return IsEmpty(RegisterAddress) AndAlso IsEmpty(RegisterGroup) AndAlso IsEmpty(RegisterName) AndAlso
               IsEmpty(D7) AndAlso IsEmpty(D6) AndAlso IsEmpty(D5) AndAlso IsEmpty(D4) AndAlso
               IsEmpty(D3) AndAlso IsEmpty(D2) AndAlso IsEmpty(D1) AndAlso IsEmpty(D0)

    End Function

    Private Shared Function IsEmpty(text As String) As Boolean
        Return String.IsNullOrWhiteSpace(text)
    End Function

End Class
