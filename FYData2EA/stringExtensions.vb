Module StringExtensions

    <System.Runtime.CompilerServices.Extension()>
    Public Function Left(ByVal input As String, ByVal length As Integer) As String

        '*** This is a simple example on how to extend the string class. ***
        '*** This method will ensure the return string is no more than x characters long. ***

        Left = input

        If String.IsNullOrEmpty(input) = False AndAlso input.Length > length Then
            Return input.Substring(0, length)
        End If

    End Function

End Module
