Imports System.Data.Common
Imports System.Runtime.Remoting

Public Class DbTransactionStub
    Inherits DbTransaction

    Public Property CommitCnt As Integer = 0
    Public Property RollbackCnt As Integer = 0
    Public Overrides ReadOnly Property IsolationLevel As IsolationLevel
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Protected Overrides ReadOnly Property DbConnection As DbConnection
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public Overrides Sub Commit()
        CommitCnt = CommitCnt + 1
    End Sub

    Public Overrides Sub Rollback()
        RollbackCnt = RollbackCnt + 1
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
End Class
