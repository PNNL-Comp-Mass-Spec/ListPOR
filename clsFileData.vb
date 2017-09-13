Public Class clsFileData
    Public ReadOnly Property Key As String
    Public ReadOnly Property Value As String
    Public ReadOnly Property ValueDbl As Double
    Public ReadOnly Property RemainingCols As String

    Public Sub New(strKey As String, strValue As String, strRemainingCols As String)
        Key = strKey
        Value = strValue
        RemainingCols = strRemainingCols

        Dim parsedValue As Double
        If Double.TryParse(strValue, parsedValue) Then
            ValueDbl = parsedValue
        Else
            ValueDbl = 0
        End If
    End Sub

    Public Overrides Function ToString() As String
        Return Key & ": " & Value
    End Function
End Class
