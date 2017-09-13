Public Class clsFileData
    Public ReadOnly Property Key As String
    Public ReadOnly Property Value As String
    Public ReadOnly Property RemainingCols As String

    Public Sub New(strKey As String, strValue As String, strRemainingCols As String)
        Key = strKey
        Value = strValue
        RemainingCols = strRemainingCols
    End Sub

    Public Overrides Function ToString() As String
        Return Key & ": " & Value
    End Function
End Class
