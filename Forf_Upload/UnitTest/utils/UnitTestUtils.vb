Imports System.Data.Common
Imports System.IO
Imports BEA.TFS.CommonUtility

Public Class UnitTestUtils
    Shared Sub BackupAndReplaceItWith(srcFile As String, destFile As String)
        File.Copy(destFile, $"{destFile}.bak", True)
        File.Copy(srcFile, destFile, True)
    End Sub

    Shared Sub CreateCleanDirectory(path As String)
        SlientDeleteDirectory(path)
        Directory.CreateDirectory(path)
    End Sub

    Shared Sub SlientDeleteDirectory(path As String)
        Try
            Directory.Delete(path, True)
        Catch ex As Exception
        End Try
    End Sub

    ''' <remarks>
    ''' Back up database data.
    ''' 
    ''' Sometimes you cannot rollback the changes done in testing (i.e. DDL changes, using sqlldr).
    ''' </remarks>
    Shared Sub CreateBackupTable(conn As IDbConnection, tableName As String, backupTableName As String)
        ExecuteNonQuery(conn, "
BEGIN
    EXECUTE IMMEDIATE 'DROP TABLE " & backupTableName & "';
EXCEPTION
    WHEN OTHERS THEN
        IF SQLCODE != -942 THEN
            RAISE;
        END IF;
END;

")
        ExecuteNonQuery(conn, "
CREATE TABLE " & backupTableName & " AS SELECT * FROM " & tableName & "
")
    End Sub

    ''' <remarks>
    ''' Restore data from backup table create by calling CreateBackupTable()
    ''' </remarks>
    Shared Sub RestoreFromBackupTable(conn As IDbConnection, tableName As String, backupTableName As String, Optional isDropBackupTable As Boolean = False)
        ExecuteNonQuery(conn, "
DELETE FROM " & tableName & "
")
        ExecuteNonQuery(conn, "
INSERT INTO " & tableName & " SELECT * FROM " & backupTableName & "
")
        If isDropBackupTable Then
            ExecuteNonQuery(conn, "
DROP TABLE " & backupTableName & "
")
        End If
    End Sub
End Class

Public Class CalledSqlExpectation
    Private Property cmds As List(Of DbCommandStub) = New List(Of DbCommandStub)
    Private Property expectedSqls As List(Of String) = New List(Of String)

    Sub New(ByRef connStub As DbConnectionStub)
        For index = 0 To connStub.CreateCommandRtnVal.Count - 1
            cmds.Add(connStub.CreateCommandRtnVal.ElementAt(index))
        Next
    End Sub

    Sub AssertEqualsTo(expectedSqls As List(Of String))
        Assert.AreEqual(expectedSqls.Count, cmds.Count)
        For index = 0 To expectedSqls.Count - 1
            Assert.AreEqual(expectedSqls(index), cmds(index).CommandTextInVal.Item(0))
        Next
    End Sub
End Class

Public Class DbParamExpectation
    Private Property expectedParamList As List(Of List(Of Object)) = New List(Of List(Of Object))
    Private Property cmds As List(Of DbCommandStub) = New List(Of DbCommandStub)

    Sub New(ByRef connStub As DbConnectionStub)
        For index = 0 To connStub.CreateCommandRtnVal.Count - 1
            cmds.Add(connStub.CreateCommandRtnVal.ElementAt(index))
        Next
    End Sub

    Sub New()
    End Sub

    Sub ClearExpectedParam()
        expectedParamList.Clear()
    End Sub

    Sub AddParam(name As String, dbType As DbType, value As Object)
        expectedParamList.Add(New List(Of Object) From {name, dbType, value})
    End Sub

    Sub AssertEqualsTo(actualParams As DbParameterCollection)
        Assert.AreEqual(expectedParamList.Count, actualParams.Count)
        For index = 0 To expectedParamList.Count - 1
            Assert.AreEqual(expectedParamList.Item(index).Item(0), actualParams.Item(index).ParameterName)
            Assert.AreEqual(expectedParamList.Item(index).Item(1), actualParams.Item(index).DbType)
            Assert.AreEqual(expectedParamList.Item(index).Item(2), actualParams.Item(index).Value)
        Next
    End Sub

    Sub AssertEqualsTo(actualParams As DbParameter())
        Assert.AreEqual(expectedParamList.Count, actualParams.Count)
        For index = 0 To expectedParamList.Count - 1
            Assert.AreEqual(expectedParamList.Item(index).Item(0), actualParams(index).ParameterName)
            Assert.AreEqual(expectedParamList.Item(index).Item(1), actualParams(index).DbType)
            Assert.AreEqual(expectedParamList.Item(index).Item(2), actualParams(index).Value)
        Next
    End Sub

    Sub AssertEqualsToCmd(index As Integer)
        Dim actualParams As DbParameterCollection = cmds(index).DbParameterCollectionInVal.ElementAt(0)
        Assert.AreEqual(expectedParamList.Count, actualParams.Count)
        For index = 0 To expectedParamList.Count - 1
            Assert.AreEqual(expectedParamList.Item(index).Item(0), actualParams(index).ParameterName)
            Assert.AreEqual(expectedParamList.Item(index).Item(1), actualParams(index).DbType)
            Assert.AreEqual(expectedParamList.Item(index).Item(2), actualParams(index).Value)
        Next
    End Sub
End Class