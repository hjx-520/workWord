Imports System.IO
Imports System.Data.OleDb
Imports BEA.TFS.CommonUtility
Imports Forf_Upload

' TODO: [REMOVE COMMENT AFTER REVIEW] Renamed class without underscore for consistency with other projects.
<TestClass()>
<DeploymentItem("..\..\res\DataSource_TEST.INI")>
<DeploymentItem("..\..\res\Forf_Upload_TEST.INI")>
Public Class FunctionTest
    Private DS_INI_FILE As String = "DataSource_TEST.INI"
    Private INI_FILE As String = "Forf_Upload_TEST.INI"

    ' TODO: Function test will not work as CR has been applied, lack of sufficient test files
    <TestMethod()> Public Sub Test_Main()
        Try
            Setup()
            Dim dsINIFile As String = Path.GetFullPath(My.MySettings.Default.DATA_SOURCE_INI_FILE)
            Dim dbConfig As DataSourceConfig = GetDataSourceConfig(dsINIFile)
            Dim rptConnStr As String = Utility.GetConnectionStr(dbConfig.RptDSN, dbConfig.RptUserID, dbConfig.RptPwd)
            Dim varINIFile As String = Path.GetFullPath(My.Settings.INI_FILE)
            Using rptConn As New OleDbConnection(rptConnStr)
                rptConn.Open()
                Forf_Upload.Main()
                Test_Happy_Forf_Upload(rptConn)
            End Using
        Finally
            Cleanup()
        End Try

    End Sub

    Private Sub Test_Happy_Forf_Upload(rptConn As IDbConnection)
        Assert.IsTrue(File.Exists("D:\TIReports\BankingReturn\Forf_Model\LoadDB\RPT_SWAP_CURVE.csv"))
        Assert.AreEqual("65268DC754B0F4962FB3F95FA1FCE77D", HashRecord("RPT_CURVE_USAGE", rptConn))
        Assert.AreEqual("BD2A03CED98B1D8436F7B99F80C34572", HashRecord("RPT_EXCPT", rptConn))
        Assert.AreEqual("896EA896A86B86231B35B7C11BD032A9", HashRecord("RPT_FORF_MASTER", rptConn))
        Assert.AreEqual("384DCBC225B08DBF625120586BDD35CA", HashRecord("RPT_FORF_UPLOAD", rptConn))
        Assert.AreEqual("74631A18F7B7D9FF0EBE5BA670A58CA8", HashRecord("RPT_IS9_PROVISION", rptConn))
    End Sub

    Private Sub Setup()
        Dim dsINIFile As String = Path.GetFullPath(My.MySettings.Default.DATA_SOURCE_INI_FILE)
        Dim dbConfig As DataSourceConfig = GetDataSourceConfig(dsINIFile)
        Dim rptConnStr As String = Utility.GetConnectionStr(dbConfig.RptDSN, dbConfig.RptUserID, dbConfig.RptPwd)

        File.Copy(INI_FILE, "Forf_Upload.INI", True)
        File.Copy("D:\TFSConfig\DataSource.INI", "D:\TFSConfig\DataSource.INI.bak", True)
        File.Copy(DS_INI_FILE, "D:\TFSConfig\DataSource.INI", True)
        Using rptConn As New OleDbConnection(rptConnStr)
            rptConn.Open()
            UnitTestUtils.CreateBackupTable(rptConn, "RPT_CURVE_USAGE", "BAK_RPT_CURVE_USAGE")
            UnitTestUtils.CreateBackupTable(rptConn, "RPT_EXCPT", "BAK_RPT_EXCPT")
            UnitTestUtils.CreateBackupTable(rptConn, "RPT_FORF_MASTER", "BAK_RPT_FORF_MASTER")
            UnitTestUtils.CreateBackupTable(rptConn, "RPT_FORF_UPLOAD", "BAK_RPT_FORF_UPLOAD")
            UnitTestUtils.CreateBackupTable(rptConn, "RPT_IS9_PROVISION", "BAK_RPT_IS9_PROVISION")
            UnitTestUtils.CreateBackupTable(rptConn, "RPT_IS9_PROVISION_STG", "BAK_RPT_IS9_PROVISION_STG")
            UnitTestUtils.CreateBackupTable(rptConn, "RPT_REPRICE_UPLOAD", "BAK_RPT_REPRICE_UPLOAD")
            UnitTestUtils.CreateBackupTable(rptConn, "RPT_SWAP_CURVE", "BAK_RPT_SWAP_CURVE")
        End Using
    End Sub

    Private Sub Cleanup()
        Dim dsINIFile As String = Path.GetFullPath(My.MySettings.Default.DATA_SOURCE_INI_FILE)
        Dim dbConfig As DataSourceConfig = GetDataSourceConfig(dsINIFile)
        Dim rptConnStr As String = Utility.GetConnectionStr(dbConfig.RptDSN, dbConfig.RptUserID, dbConfig.RptPwd)

        File.Copy("D:\TFSConfig\DataSource.INI.bak", "D:\TFSConfig\DataSource.INI", True)
        Using rptConn As New OleDbConnection(rptConnStr)
            rptConn.Open()
            UnitTestUtils.RestoreFromBackupTable(rptConn, "RPT_CURVE_USAGE", "BAK_RPT_CURVE_USAGE")
            UnitTestUtils.RestoreFromBackupTable(rptConn, "RPT_EXCPT", "BAK_RPT_EXCPT")
            UnitTestUtils.RestoreFromBackupTable(rptConn, "RPT_FORF_MASTER", "BAK_RPT_FORF_MASTER")
            UnitTestUtils.RestoreFromBackupTable(rptConn, "RPT_FORF_UPLOAD", "BAK_RPT_FORF_UPLOAD")
            UnitTestUtils.RestoreFromBackupTable(rptConn, "RPT_IS9_PROVISION", "BAK_RPT_IS9_PROVISION")
            UnitTestUtils.RestoreFromBackupTable(rptConn, "RPT_IS9_PROVISION_STG", "BAK_RPT_IS9_PROVISION_STG")
            UnitTestUtils.RestoreFromBackupTable(rptConn, "RPT_REPRICE_UPLOAD", "BAK_RPT_REPRICE_UPLOAD")
            UnitTestUtils.RestoreFromBackupTable(rptConn, "RPT_SWAP_CURVE", "BAK_RPT_SWAP_CURVE")
        End Using
    End Sub

    <TestMethod()> Public Sub Test_ValidateForfaiting()
        Setup()
        Dim dsINIFile As String = Path.GetFullPath(My.MySettings.Default.DATA_SOURCE_INI_FILE)
        Dim dbConfig As DataSourceConfig = GetDataSourceConfig(dsINIFile)
        Dim rptConnStr As String = Utility.GetConnectionStr(dbConfig.RptDSN, dbConfig.RptUserID, dbConfig.RptPwd)
        Dim iniFile As String = Path.GetFullPath(My.MySettings.Default.INI_FILE)
        Dim config As IniConfig = GetConfig(iniFile)
        config = GetConfig(iniFile)
        Dim beforeCountRPT_EXCPT As Integer
        Dim afterCountRPT_EXCPT As Integer

        Try
            Using rptConn As New OleDbConnection(rptConnStr)
                rptConn.Open()
                Using rptTran As OleDbTransaction = rptConn.BeginTransaction()
                    beforeCountRPT_EXCPT = Test_SQL_Count_RPT_EXCPT(rptConn, rptTran)
                    Utility.ValidateForfaiting(rptConn, rptTran)
                    afterCountRPT_EXCPT = Test_SQL_Count_RPT_EXCPT(rptConn, rptTran)
                End Using
            End Using
            Assert.AreEqual(0, afterCountRPT_EXCPT)
            Assert.AreEqual(1, beforeCountRPT_EXCPT)
        Finally
            Cleanup()
        End Try
    End Sub

    <TestMethod()> Public Sub Test_ValidaterReprice()
        Setup()
        Dim dsINIFile As String = Path.GetFullPath(My.MySettings.Default.DATA_SOURCE_INI_FILE)
        Dim dbConfig As DataSourceConfig = GetDataSourceConfig(dsINIFile)
        Dim rptConnStr As String = Utility.GetConnectionStr(dbConfig.RptDSN, dbConfig.RptUserID, dbConfig.RptPwd)
        Dim iniFile As String = Path.GetFullPath(My.MySettings.Default.INI_FILE)
        Dim config As IniConfig = GetConfig(iniFile)
        config = GetConfig(iniFile)
        Dim beforeCountRPT_EXCPT As Integer
        Dim afterCountRPT_EXCPT As Integer

        Try
            Using rptConn As New OleDbConnection(rptConnStr)
                rptConn.Open()
                Using rptTran As OleDbTransaction = rptConn.BeginTransaction()
                    beforeCountRPT_EXCPT = Test_SQL_Count_RPT_EXCPT(rptConn, rptTran)
                    Utility.ValidateReprice(rptConn, rptTran)
                    afterCountRPT_EXCPT = Test_SQL_Count_RPT_EXCPT(rptConn, rptTran)
                End Using
            End Using
            Assert.AreEqual(1, afterCountRPT_EXCPT)
            Assert.AreEqual(1, beforeCountRPT_EXCPT)
        Finally
            Cleanup()
        End Try
    End Sub

    <TestMethod()> Public Sub Test_ValidaterRepriceForOthers()
        Setup()
        Dim dsINIFile As String = Path.GetFullPath(My.MySettings.Default.DATA_SOURCE_INI_FILE)
        Dim dbConfig As DataSourceConfig = GetDataSourceConfig(dsINIFile)
        Dim rptConnStr As String = Utility.GetConnectionStr(dbConfig.RptDSN, dbConfig.RptUserID, dbConfig.RptPwd)
        Dim iniFile As String = Path.GetFullPath(My.MySettings.Default.INI_FILE)
        Dim config As IniConfig = GetConfig(iniFile)
        config = GetConfig(iniFile)
        Dim beforeCountRPT_EXCPT As Integer
        Dim afterCountRPT_EXCPT As Integer

        Try
            Using rptConn As New OleDbConnection(rptConnStr)
                rptConn.Open()
                Using rptTran As OleDbTransaction = rptConn.BeginTransaction()
                    beforeCountRPT_EXCPT = Test_SQL_Count_RPT_EXCPT(rptConn, rptTran)
                    Utility.ValidateRepriceForOthers(rptConn, rptTran)
                    afterCountRPT_EXCPT = Test_SQL_Count_RPT_EXCPT(rptConn, rptTran)
                End Using
            End Using
            Assert.AreEqual(1, afterCountRPT_EXCPT)
            Assert.AreEqual(1, beforeCountRPT_EXCPT)
        Finally
            Cleanup()
        End Try
    End Sub

    <TestMethod()> Public Sub Test_ValidaterRepriceForPeriod()
        Setup()
        Dim dsINIFile As String = Path.GetFullPath(My.MySettings.Default.DATA_SOURCE_INI_FILE)
        Dim dbConfig As DataSourceConfig = GetDataSourceConfig(dsINIFile)
        Dim rptConnStr As String = Utility.GetConnectionStr(dbConfig.RptDSN, dbConfig.RptUserID, dbConfig.RptPwd)
        Dim iniFile As String = Path.GetFullPath(My.MySettings.Default.INI_FILE)
        Dim config As IniConfig = GetConfig(iniFile)
        config = GetConfig(iniFile)
        Dim beforeCountRPT_EXCPT As Integer
        Dim afterCountRPT_EXCPT As Integer

        Try
            Using rptConn As New OleDbConnection(rptConnStr)
                rptConn.Open()
                Using rptTran As OleDbTransaction = rptConn.BeginTransaction()
                    beforeCountRPT_EXCPT = Test_SQL_Count_RPT_EXCPT(rptConn, rptTran)
                    Utility.ValidateRepriceForPeriod(rptConn, rptTran)
                    afterCountRPT_EXCPT = Test_SQL_Count_RPT_EXCPT(rptConn, rptTran)
                End Using
            End Using
            Assert.AreEqual(3, afterCountRPT_EXCPT)
            Assert.AreEqual(1, beforeCountRPT_EXCPT)
        Finally
            Cleanup()
        End Try
    End Sub

    <TestMethod()> Public Sub Test_ValidateFullSet()
        Setup()
        Dim dsINIFile As String = Path.GetFullPath(My.MySettings.Default.DATA_SOURCE_INI_FILE)
        Dim dbConfig As DataSourceConfig = GetDataSourceConfig(dsINIFile)
        Dim rptConnStr As String = Utility.GetConnectionStr(dbConfig.RptDSN, dbConfig.RptUserID, dbConfig.RptPwd)
        Dim iniFile As String = Path.GetFullPath(My.MySettings.Default.INI_FILE)
        Dim config As IniConfig = GetConfig(iniFile)
        config = GetConfig(iniFile)
        Dim beforeCountRPT_EXCPT As Integer
        Dim afterCountRPT_EXCPT As Integer

        Try
            Using rptConn As New OleDbConnection(rptConnStr)
                rptConn.Open()
                Using rptTran As OleDbTransaction = rptConn.BeginTransaction()
                    beforeCountRPT_EXCPT = Test_SQL_Count_RPT_EXCPT(rptConn, rptTran)
                    Utility.ValidateFullSet(rptConn, rptTran)
                    afterCountRPT_EXCPT = Test_SQL_Count_RPT_EXCPT(rptConn, rptTran)
                End Using
            End Using
            Assert.AreEqual(1, afterCountRPT_EXCPT)
            Assert.AreEqual(1, beforeCountRPT_EXCPT)
        Finally
            Cleanup()
        End Try
    End Sub

    <TestMethod()> Public Sub Test_ValidateCurve()
        Setup()
        Dim dsINIFile As String = Path.GetFullPath(My.MySettings.Default.DATA_SOURCE_INI_FILE)
        Dim dbConfig As DataSourceConfig = GetDataSourceConfig(dsINIFile)
        Dim rptConnStr As String = Utility.GetConnectionStr(dbConfig.RptDSN, dbConfig.RptUserID, dbConfig.RptPwd)
        Dim iniFile As String = Path.GetFullPath(My.MySettings.Default.INI_FILE)
        Dim config As IniConfig = GetConfig(iniFile)
        config = GetConfig(iniFile)
        Dim beforeCountRPT_EXCPT As Integer
        Dim afterCountRPT_EXCPT As Integer

        Try
            Using rptConn As New OleDbConnection(rptConnStr)
                rptConn.Open()
                Using rptTran As OleDbTransaction = rptConn.BeginTransaction()
                    beforeCountRPT_EXCPT = Test_SQL_Count_RPT_EXCPT(rptConn, rptTran)
                    Utility.ValidateCurve(rptConn, rptTran)
                    afterCountRPT_EXCPT = Test_SQL_Count_RPT_EXCPT(rptConn, rptTran)
                End Using
            End Using
            Assert.AreEqual(1, afterCountRPT_EXCPT)
            Assert.AreEqual(1, beforeCountRPT_EXCPT)
        Finally
            Cleanup()
        End Try
    End Sub

    <TestMethod()> Public Sub Test_InsertForfaitingMaster()
        Setup()
        Dim dsINIFile As String = Path.GetFullPath(My.MySettings.Default.DATA_SOURCE_INI_FILE)
        Dim dbConfig As DataSourceConfig = GetDataSourceConfig(dsINIFile)
        Dim rptConnStr As String = Utility.GetConnectionStr(dbConfig.RptDSN, dbConfig.RptUserID, dbConfig.RptPwd)
        Dim iniFile As String = Path.GetFullPath(My.MySettings.Default.INI_FILE)
        Dim config As IniConfig = GetConfig(iniFile)
        config = GetConfig(iniFile)
        Dim beforeCountRPT_FORF_MASTER As Integer
        Dim afterCountRPT_FORF_MASTER As Integer

        Try
            Using rptConn As New OleDbConnection(rptConnStr)
                rptConn.Open()
                Using rptTran As OleDbTransaction = rptConn.BeginTransaction()
                    beforeCountRPT_FORF_MASTER = Test_SQL_Count_RPT_FORF_MASTER(rptConn, rptTran)
                    Utility.InsertForfaitingMaster(rptConn, rptTran)
                    afterCountRPT_FORF_MASTER = Test_SQL_Count_RPT_FORF_MASTER(rptConn, rptTran)
                End Using
            End Using
            Assert.AreEqual(184, afterCountRPT_FORF_MASTER)
            Assert.AreEqual(184, beforeCountRPT_FORF_MASTER)
        Finally
            Cleanup()
        End Try
    End Sub
    Private Function Test_SQL_Count_RPT_EXCPT(conn As IDbConnection, tran As IDbTransaction) As Integer
        Dim dsINIFile As String = Path.GetFullPath(My.MySettings.Default.DATA_SOURCE_INI_FILE)
        Dim dbConfig As DataSourceConfig = GetDataSourceConfig(dsINIFile)
        Dim rptConnStr As String = Utility.GetConnectionStr(dbConfig.RptDSN, dbConfig.RptUserID, dbConfig.RptPwd)
        Dim count As Integer

        Using cmd = conn.CreateCommand()
            cmd.Connection = conn
            cmd.Transaction = tran
            cmd.CommandText = "SELECT COUNT(*) FROM RPT_EXCPT"
            count = Convert.ToInt32(cmd.ExecuteScalar)
        End Using
        Return count
    End Function
    Private Function Test_SQL_Count_RPT_FORF_MASTER(conn As IDbConnection, tran As IDbTransaction) As Integer
        Dim dsINIFile As String = Path.GetFullPath(My.MySettings.Default.DATA_SOURCE_INI_FILE)
        Dim dbConfig As DataSourceConfig = GetDataSourceConfig(dsINIFile)
        Dim rptConnStr As String = Utility.GetConnectionStr(dbConfig.RptDSN, dbConfig.RptUserID, dbConfig.RptPwd)
        Dim count As Integer

        Using cmd = conn.CreateCommand()
            cmd.Connection = conn
            cmd.Transaction = tran
            cmd.CommandText = "SELECT COUNT(*) FROM RPT_FORF_MASTER"
            count = Convert.ToInt32(cmd.ExecuteScalar)
        End Using
        Return count
    End Function



End Class