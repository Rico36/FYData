'Come on now, they should be on
Option Strict On
Option Explicit On
Public Class cTally

    'Function name, no of records, string
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
    Private _data As String
    Public Property Data() As String
        Get
            Return _data
        End Get
        Set(ByVal value As String)
            _data = value
        End Set
    End Property

End Class

