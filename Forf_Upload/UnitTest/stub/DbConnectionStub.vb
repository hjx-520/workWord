Imports System.ComponentModel
Imports System.Data.Common
Imports System.Runtime.Remoting
Imports System.Threading

Public Class DbConnectionStub
    Inherits DbConnection
    Implements ICloneable, IDbConnection, IDisposable

    Public Property CreateCommandRtnVal As Queue(Of DbCommand) = New Queue(Of DbCommand)
    Public Property OpenCnt As Integer = 0
    Public Property CloseCnt As Integer = 0

    Public Overrides Property Site As ISite
        Get
            Return MyBase.Site
        End Get
        Set(value As ISite)
            MyBase.Site = value
        End Set
    End Property

    Public Overrides Property ConnectionString As String
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As String)
            Throw New NotImplementedException()
        End Set
    End Property

    Public Overrides ReadOnly Property ConnectionTimeout As Integer
        Get
            Return MyBase.ConnectionTimeout
        End Get
    End Property

    Public Overrides ReadOnly Property Database As String
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public Overrides ReadOnly Property DataSource As String
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public Overrides ReadOnly Property ServerVersion As String
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public Overrides ReadOnly Property State As ConnectionState
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Protected Overrides ReadOnly Property CanRaiseEvents As Boolean
        Get
            Return MyBase.CanRaiseEvents
        End Get
    End Property

    Protected Overrides ReadOnly Property DbProviderFactory As DbProviderFactory
        Get
            Return MyBase.DbProviderFactory
        End Get
    End Property

    Public Overrides Sub Close()
        CloseCnt = CloseCnt + 1
    End Sub

    Public Overrides Sub ChangeDatabase(databaseName As String)
        Throw New NotImplementedException()
    End Sub

    Public Overrides Sub Open()
        OpenCnt = OpenCnt + 1
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        MyBase.Dispose(disposing)
    End Sub

    Protected Overrides Sub OnStateChange(stateChange As StateChangeEventArgs)
        MyBase.OnStateChange(stateChange)
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

    Public Overrides Function GetSchema() As DataTable
        Return MyBase.GetSchema()
    End Function

    Public Overrides Function GetSchema(collectionName As String) As DataTable
        Return MyBase.GetSchema(collectionName)
    End Function

    Public Overrides Function GetSchema(collectionName As String, restrictionValues() As String) As DataTable
        Return MyBase.GetSchema(collectionName, restrictionValues)
    End Function

    Public Overrides Function OpenAsync(cancellationToken As CancellationToken) As Task
        Return MyBase.OpenAsync(cancellationToken)
    End Function

    Public Function Clone() As Object Implements ICloneable.Clone
        Throw New NotImplementedException()
    End Function

    Protected Overrides Function GetService(service As Type) As Object
        Return MyBase.GetService(service)
    End Function

    Protected Overrides Function BeginDbTransaction(isolationLevel As IsolationLevel) As DbTransaction
        Return New DbTransactionStub()
    End Function

    Protected Overrides Function CreateDbCommand() As DbCommand
        If CreateCommandRtnVal.Count > 0 Then
            Return CreateCommandRtnVal.Dequeue()
        End If
        Return Nothing
    End Function

End Class

