Imports System.Data.Common
Imports System.IO
Imports System.Runtime.Remoting
Imports System.Threading

Public Class DbDataReaderStub
    Inherits DbDataReader

    Public MockFields As Dictionary(Of String, Type) = New Dictionary(Of String, Type)
    Public MockRecords As List(Of List(Of Object)) = New List(Of List(Of Object))
    Private MockRecordIndex As Integer = -1
    Public Overrides ReadOnly Property Depth As Integer
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public Overrides ReadOnly Property FieldCount As Integer
        Get
            Return MockFields.Count()
        End Get
    End Property

    Public Overrides ReadOnly Property HasRows As Boolean
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public Overrides ReadOnly Property IsClosed As Boolean
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public Overrides ReadOnly Property RecordsAffected As Integer
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public Overrides ReadOnly Property VisibleFieldCount As Integer
        Get
            Return MyBase.VisibleFieldCount
        End Get
    End Property

    Default Public Overrides ReadOnly Property Item(ordinal As Integer) As Object
        Get
            Return MockRecords.Item(MockRecordIndex).Item(ordinal)
        End Get
    End Property

    Default Public Overrides ReadOnly Property Item(name As String) As Object
        Get
            Return MockRecords.Item(MockRecordIndex).Item(GetOrdinal(name))
        End Get
    End Property

    Public Overrides Sub Close()
        MyBase.Close()
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        MyBase.Dispose(disposing)
    End Sub

    Public Overrides Function ToString() As String
        Return MyBase.ToString()
    End Function

    Public Overrides Function Equals(obj As Object) As Boolean
        Return MyBase.Equals(obj)
    End Function

    Public Overrides Function GetHashCode() As Integer
        Return MyBase.GetHashCode()
    End Function

    Public Overrides Function InitializeLifetimeService() As Object
        Return MyBase.InitializeLifetimeService()
    End Function

    Public Overrides Function CreateObjRef(requestedType As Type) As ObjRef
        Return MyBase.CreateObjRef(requestedType)
    End Function

    Public Overrides Function GetDataTypeName(ordinal As Integer) As String
        Throw New NotImplementedException()
    End Function

    Public Overrides Function GetEnumerator() As IEnumerator
        Throw New NotImplementedException()
    End Function

    Public Overrides Function GetFieldType(ordinal As Integer) As Type
        Return MockFields.Values(ordinal)
    End Function

    Public Overrides Function GetName(ordinal As Integer) As String
        Return MockFields.Keys(ordinal)
    End Function

    Public Overrides Function GetOrdinal(name As String) As Integer
        Return MockFields.ToList.IndexOf(MockFields.First(Function(x) x.Key = name))
    End Function

    Public Overrides Function GetSchemaTable() As DataTable
        Return MyBase.GetSchemaTable()
    End Function

    Public Overrides Function GetBoolean(ordinal As Integer) As Boolean
        Return Item(ordinal)
    End Function

    Public Overrides Function GetByte(ordinal As Integer) As Byte
        Return Item(ordinal)
    End Function

    Public Overrides Function GetBytes(ordinal As Integer, dataOffset As Long, buffer() As Byte, bufferOffset As Integer, length As Integer) As Long
        Return Item(ordinal)
    End Function

    Public Overrides Function GetChar(ordinal As Integer) As Char
        Return Item(ordinal)
    End Function

    Public Overrides Function GetChars(ordinal As Integer, dataOffset As Long, buffer() As Char, bufferOffset As Integer, length As Integer) As Long
        Return Item(ordinal)
    End Function

    Public Overrides Function GetDateTime(ordinal As Integer) As Date
        Return Item(ordinal)
    End Function

    Public Overrides Function GetDecimal(ordinal As Integer) As Decimal
        Return Item(ordinal)
    End Function

    Public Overrides Function GetDouble(ordinal As Integer) As Double
        Return Item(ordinal)
    End Function

    Public Overrides Function GetFloat(ordinal As Integer) As Single
        Return Item(ordinal)
    End Function

    Public Overrides Function GetGuid(ordinal As Integer) As Guid
        Return Item(ordinal)
    End Function

    Public Overrides Function GetInt16(ordinal As Integer) As Short
        Return Item(ordinal)
    End Function

    Public Overrides Function GetInt32(ordinal As Integer) As Integer
        Return Item(ordinal)
    End Function

    Public Overrides Function GetInt64(ordinal As Integer) As Long
        Return Item(ordinal)
    End Function

    Public Overrides Function GetProviderSpecificFieldType(ordinal As Integer) As Type
        Return MyBase.GetProviderSpecificFieldType(ordinal)
    End Function

    Public Overrides Function GetProviderSpecificValue(ordinal As Integer) As Object
        Return MyBase.GetProviderSpecificValue(ordinal)
    End Function

    Public Overrides Function GetProviderSpecificValues(values() As Object) As Integer
        Return MyBase.GetProviderSpecificValues(values)
    End Function

    Public Overrides Function GetString(ordinal As Integer) As String
        Return Item(ordinal)
    End Function

    Public Overrides Function GetStream(ordinal As Integer) As Stream
        Return MyBase.GetStream(ordinal)
    End Function

    Public Overrides Function GetTextReader(ordinal As Integer) As TextReader
        Return MyBase.GetTextReader(ordinal)
    End Function

    Public Overrides Function GetValue(ordinal As Integer) As Object
        Return Item(ordinal)
    End Function

    Public Overrides Function GetFieldValue(Of T)(ordinal As Integer) As T
        Return MyBase.GetFieldValue(Of T)(ordinal)
    End Function

    Public Overrides Function GetFieldValueAsync(Of T)(ordinal As Integer, cancellationToken As CancellationToken) As Task(Of T)
        Return MyBase.GetFieldValueAsync(Of T)(ordinal, cancellationToken)
    End Function

    Public Overrides Function GetValues(values() As Object) As Integer
        Throw New NotImplementedException()
    End Function

    Public Overrides Function IsDBNull(ordinal As Integer) As Boolean
        Return Information.IsDBNull(Item(ordinal))
    End Function

    Public Overrides Function IsDBNullAsync(ordinal As Integer, cancellationToken As CancellationToken) As Task(Of Boolean)
        Return MyBase.IsDBNullAsync(ordinal, cancellationToken)
    End Function

    Public Overrides Function NextResult() As Boolean
        Throw New NotImplementedException()
    End Function

    Public Overrides Function Read() As Boolean
        If MockRecordIndex < MockRecords.Count() - 1 Then
            MockRecordIndex = MockRecordIndex + 1
            Return True
        End If
        Return False
    End Function

    Public Overrides Function ReadAsync(cancellationToken As CancellationToken) As Task(Of Boolean)
        Return MyBase.ReadAsync(cancellationToken)
    End Function

    Public Overrides Function NextResultAsync(cancellationToken As CancellationToken) As Task(Of Boolean)
        Return MyBase.NextResultAsync(cancellationToken)
    End Function

    Protected Overrides Function GetDbDataReader(ordinal As Integer) As DbDataReader
        Return MyBase.GetDbDataReader(ordinal)
    End Function
End Class
