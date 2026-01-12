Imports CommonUtilities
Imports BEA.TFS.CommonUtility
Imports GenMA23Report

Public Class ConnectionStubBuilder
    Private Property cmdList As List(Of DbCommandStub) = New List(Of DbCommandStub)
    Private Property readerList As List(Of DbDataReaderStub) = New List(Of DbDataReaderStub)

    Public Shared Function NewInstance() As ConnectionStubBuilder
        Return New ConnectionStubBuilder()
    End Function

    Public Function AddQueryRtnVal() As ConnectionStubBuilder
        cmdList.Add(New DbCommandStub())
        readerList.Add(New DbDataReaderStub())
        Return Me
    End Function

    Public Function AddExceptionOnQuery() As ConnectionStubBuilder
        cmdList.Add(Nothing)
        Return Me
    End Function

    Public Function WithFields(ByVal fields As Dictionary(Of String, Type)) As ConnectionStubBuilder
        readerList.Last.MockFields = fields
        Return Me
    End Function

    Public Function AddRow(ByVal row As List(Of Object)) As ConnectionStubBuilder
        readerList.Last.MockRecords.Add(row)
        Return Me
    End Function

    Public Function Build() As DbConnectionStub
        Dim artifact As DbConnectionStub = New DbConnectionStub()
        For index = 0 To cmdList.Count - 1
            Dim currentCmd As DbCommandStub = cmdList.Item(index)
            If currentCmd IsNot Nothing Then
                currentCmd.ExecuteReaderRtnVal.Enqueue(readerList.Item(index))
            End If
            artifact.CreateCommandRtnVal.Enqueue(currentCmd)
        Next
        Return artifact
    End Function
End Class

Public Class ResultSetStubBuilder
    Private Property reader As DbDataReaderStub = New DbDataReaderStub()

    Public Shared Function NewInstance() As ResultSetStubBuilder
        Return New ResultSetStubBuilder()
    End Function

    Public Function WithFields(ByVal fields As Dictionary(Of String, Type)) As ResultSetStubBuilder
        reader.MockFields = fields
        Return Me
    End Function

    Public Function AddRow(ByVal row As List(Of Object)) As ResultSetStubBuilder
        reader.MockRecords.Add(row)
        Return Me
    End Function

    Public Function Build() As DbResultSet
        Return New DbResultSet(reader)
    End Function
End Class