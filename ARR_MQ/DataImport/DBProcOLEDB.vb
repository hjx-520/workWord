Imports System.Data
Imports System.Data.OleDb

Public Class DBProcOLEDB
    Public objconn As OleDb.OleDbConnection
    Public objda As OleDb.OleDbDataAdapter
    Public objcmd As OleDb.OleDbCommand
    Public strConnString As String

    Public Function GetConnectionStr(dataSource As String, userID As String, password As String) As String
        strConnString = $"Provider=OraOLEDB.Oracle;Data Source={dataSource};User Id={userID};Password={password};Persist Security Info=False"
        'strConnString = $"Provider=OraOLEDB.Oracle; Host=hkgapvt-tsd61:11521; Initial Catalog=TIIUUATR; User Id={userID};Password={password};"
        Return strConnString
    End Function

    Public Function DBOpen() As Boolean
        'Dim strConnString As String = "Data Source=FTRDEV;Persist Security Info=True;User ID=TRADEIN1;Password=TRADEIN1;Load Balance Timeout=1000"
        'Dim strConnString As String = "server=hkgfpkt03; database=ELP; user id=sa; password=; Connect Timeout=15"
        objconn = New OleDbConnection(strConnString)
        objcmd = New OleDbCommand
        objda = New OleDbDataAdapter(objcmd)
        objcmd.Connection = objconn

        Try
            objconn.Open()
            Return True
        Catch ex As Exception

            Console.WriteLine("Connect to database failed! with Connection String is " + strConnString + " " + ex.Message)
            Return False

        End Try
    End Function

    'Close database
    Public Function Dispose() As Boolean
        Try
            objconn.Close()
            objconn.Dispose()
            objcmd.Dispose()
            objda.Dispose()
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Sub CloseConnection()
        objconn.Close()
    End Sub

    'return a dataset by execute sql query
    Public Overloads Function FillDataSet(ByVal strSQL As String) As DataSet
        Dim ds As DataSet = New DataSet
        If objconn.State = ConnectionState.Closed Then
            objconn.Open()
        End If
        Try
            objcmd.CommandText = strSQL
            objda.Fill(ds, "Table")
        Catch ex As Exception
            Return Nothing
        End Try

        Return ds
    End Function
    'return a dataset that create table by table name and by execute sql query
    Public Overloads Function FillDataSet(ByVal strSQL As String, ByVal strTableName As String) As DataSet
        If objconn.State = ConnectionState.Closed Then
            objconn.Open()
        End If
        Dim ds As DataSet = New DataSet
        Try
            objcmd.CommandText = strSQL
            objda.Fill(ds, strTableName)
        Catch ex As Exception
            Return Nothing
        End Try
        Return ds
    End Function

    ' execute sql return
    Public Overloads Function ExecuteSQL(ByVal strSQL As String) As Boolean
        If objconn.State = ConnectionState.Closed Then
            objconn.Open()
        End If

        objcmd.CommandText = strSQL
        Try
            objcmd.ExecuteNonQuery()
            Return True
        Catch ex As Exception

            Return False
        Finally

        End Try
    End Function

    'return a sql reader by executing sql query
    Public Function ExecuteSQLReader(ByVal strSQL As String) As OleDb.OleDbDataReader

        Dim result As OleDbDataReader
        If objconn.State = ConnectionState.Closed Then
            objconn.Open()
        End If
        objcmd.CommandText = strSQL
        Try
            result = objcmd.ExecuteReader
        Catch ex As Exception

            Return Nothing
        End Try

        Return result
    End Function
    'return short number to count records in table
    Public Function ExecuteSQLScalar(ByVal strSQL As String) As Short
        If objconn.State = ConnectionState.Closed Then
            objconn.Open()
        End If
        objcmd.CommandText = strSQL
        Dim iCount As Short = 0
        Try
            iCount = Int(objcmd.ExecuteScalar)
        Catch ex As Exception
            objconn.Close()
            Return 0
        Finally

        End Try
        Return iCount

    End Function
    Public Function ExecuteSQLTrans(ByVal strSQL() As String) As Boolean

        If objconn.State = ConnectionState.Closed Then
            objconn.Open()
        End If

        Dim myTrans As OleDb.OleDbTransaction

        ' Start a local transaction
        myTrans = objconn.BeginTransaction(IsolationLevel.ReadCommitted)
        ' Assign transaction object for a pending local transaction
        objcmd.Transaction = myTrans

        Try
            For Each strCmd As String In strSQL
                objcmd.CommandText = strCmd
                objcmd.ExecuteNonQuery()
            Next

            myTrans.Commit()
            Return False
        Catch e As Exception
            myTrans.Rollback()
            objconn.Close()
            Return False

        End Try

    End Function
End Class
