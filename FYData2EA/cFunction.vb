'Come on now, they should be on
Option Strict On
Option Explicit On
Public Class cFunction

    'Function name, no of records in ESC, FYMinus2, FYMinus1, FY
    Private _name As String
        Public Property Name() As String
            Get
                Return _name
            End Get
            Set(ByVal value As String)
                _name = value
            End Set
        End Property

    Private _num As Integer
    Public Property num() As Integer
        Get
            Return _num
        End Get
        Set(ByVal value As Integer)
            _num = value
        End Set
    End Property

    Private _fyminus2 As Double
    Public Property fyminus2() As Double
        Get
            Return _fyminus2
        End Get
        Set(ByVal value As Double)
            _fyminus2 = value
        End Set
    End Property
    Private _fyminus1 As Double
    Public Property fyminus1() As Double
        Get
            Return _fyminus1
        End Get
        Set(ByVal value As Double)
            _fyminus1 = value
        End Set
    End Property
    Private _fy As Double
    Public Property fy() As Double
        Get
            Return _fy
        End Get
        Set(ByVal value As Double)
            _fy = value
        End Set
    End Property

End Class
