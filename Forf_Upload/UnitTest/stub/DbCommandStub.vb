Imports System.ComponentModel
Imports System.Data.Common
Imports System.Data.OleDb
Imports System.Runtime.Remoting
Imports System.Threading

Public Class DbCommandStub
    Inherits DbCommand
    Implements ICloneable, IDbCommand, IDisposable

    Public Property CommandTextInVal As List(Of String) = New List(Of String)
    Public Property TransactionInVal As List(Of DbTransaction) = New List(Of DbTransaction)
    Public Property DbParameterCollectionInVal As List(Of DbParameterCollection) = New List(Of DbParameterCollection) From {(New OleDbCommand).Parameters}
    Public Property ExecuteNonQueryRtnVal As Queue(Of Integer) = New Queue(Of Integer)
    Public Property ExecuteReaderRtnVal As Queue(Of DbDataReader) = New Queue(Of DbDataReader)
    Public Overrides Property Site As ISite
        Get
            Return MyBase.Site
        End Get
        Set(value As ISite)
            MyBase.Site = value
        End Set
    End Property

    Public Overrides Property CommandText As String
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As String)
            CommandTextInVal.Add(value)
        End Set
    End Property

    Public Overrides Property CommandTimeout As Integer
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As Integer)
            Throw New NotImplementedException()
        End Set
    End Property

    Public Overrides Property CommandType As CommandType
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As CommandType)
            Throw New NotImplementedException()
        End Set
    End Property

    Public Overrides Property DesignTimeVisible As Boolean
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As Boolean)
            Throw New NotImplementedException()
        End Set
    End Property

    Public Overrides Property UpdatedRowSource As UpdateRowSource
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As UpdateRowSource)
            Throw New NotImplementedException()
        End Set
    End Property

    Protected Overrides ReadOnly Property CanRaiseEvents As Boolean
        Get
            Return MyBase.CanRaiseEvents
        End Get
    End Property

    Protected Overrides Property DbConnection As DbConnection
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As DbConnection)
            Throw New NotImplementedException()
        End Set
    End Property

    Protected Overrides ReadOnly Property DbParameterCollection As DbParameterCollection
        Get
            Return DbParameterCollectionInVal.First
        End Get
    End Property

    Protected Overrides Property DbTransaction As DbTransaction
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As DbTransaction)
            TransactionInVal.Add(value)
        End Set
    End Property

    Public Overrides Sub Cancel()
        Throw New NotImplementedException()
    End Sub

    Public Overrides Sub Prepare()
        Throw New NotImplementedException()
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        MyBase.Dispose(disposing)
    End Sub

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

    Public Overrides Function ToString() As String
        Return MyBase.ToString()
    End Function

    Public Overrides Function ExecuteNonQuery() As Integer
        If ExecuteNonQueryRtnVal.Count > 0 Then
            Return ExecuteNonQueryRtnVal.Dequeue()
        End If
        Return Nothing
    End Function

    Public Overrides Function ExecuteNonQueryAsync(cancellationToken As CancellationToken) As Task(Of Integer)
        Return MyBase.ExecuteNonQueryAsync(cancellationToken)
    End Function

    Public Overrides Function ExecuteScalarAsync(cancellationToken As CancellationToken) As Task(Of Object)
        Return MyBase.ExecuteScalarAsync(cancellationToken)
    End Function

    Public Overrides Function ExecuteScalar() As Object
        Dim r = ExecuteDbDataReader(CommandBehavior.SingleResult)
        If r.Read() Then
            Return r.GetValue(0)
        Else
            Return Nothing
        End If
    End Function

    Public Function Clone() As Object Implements ICloneable.Clone
        Throw New NotImplementedException()
    End Function

    Protected Overrides Function GetService(service As Type) As Object
        Return MyBase.GetService(service)
    End Function

    Protected Overrides Function CreateDbParameter() As DbParameter
        Return New OleDbParameter()
    End Function

    Protected Overrides Function ExecuteDbDataReader(behavior As CommandBehavior) As DbDataReader
        If ExecuteReaderRtnVal.Count > 0 Then
            Return ExecuteReaderRtnVal.Dequeue()
        End If
        Return New DbDataReaderStub()
    End Function

    Protected Overrides Function ExecuteDbDataReaderAsync(behavior As CommandBehavior, cancellationToken As CancellationToken) As Task(Of DbDataReader)
        Return MyBase.ExecuteDbDataReaderAsync(behavior, cancellationToken)
    End Function
End Class
