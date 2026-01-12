Public Class ImportFields

    Public Property FileName As String
    Public Property TableName As String
    Public Property FileHeader As String
    Public Property FileTrailer As String
    Public Property FieldNum As Integer
    Public Property FieldName() As New List(Of String)
    Public Property DataType() As New List(Of String)
    Public Property StartPos() As New List(Of Integer)
    Public Property FieldLength() As New List(Of Integer)

    Public Property FieldSource() As New List(Of String)

End Class
