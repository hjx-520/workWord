Imports System.IO
Imports log4net
Imports Newtonsoft.Json
Imports BEA.TFS.CommonUtility
Imports BEA.TFS.CommonUtility.CommonEntityUtilities
Imports System.Data.OleDb
Imports System.Collections.ObjectModel

Module Forf_Upload
    Private ReadOnly _logger As ILog = LogManager.GetLogger(GetType(Forf_Upload).ToString())
    Sub Main()
        Try
            _logger.Info("Start Forf_Upload...")
            _logger.Info("Initializing...")

            Dim dsINIFile As String = Path.GetFullPath(My.MySettings.Default.DATA_SOURCE_INI_FILE)
            Dim dbConfig As DataSourceConfig = GetDataSourceConfig(dsINIFile)
            Dim iniFile As String = Path.GetFullPath(My.MySettings.Default.INI_FILE)
            Dim config As IniConfig = GetConfig(iniFile)
            Dim flagINIFile As String = Path.GetFullPath(My.MySettings.Default.FLAG_FILE_INI_FILE)
            Dim flagFile As String = GetProfileStringVal("FLAGFILE", "Forf_Upload", flagINIFile)
            Dim flagFileForfGenRepaySch As String = GetProfileStringVal("FLAGFILE", "Forf_GenRepaySch", flagINIFile)
            Dim rptConnStr As String = Utility.GetConnectionStr(dbConfig.RptDSN, dbConfig.RptUserID, dbConfig.RptPwd)


            Dim strKPS_ref As String = "KPS_CURVE"
            Dim strTCL_ref As String = "TCL_CURVE"

            If Not File.Exists(flagFile) Then
                File.Create(flagFile).Close()
            End If

            Using conn As IDbConnection = New OleDbConnection(rptConnStr)
                conn.Open()
                Using tran As IDbTransaction = conn.BeginTransaction()
                    Try
                        _logger.Info("Get Dates...")
                        Dim rptDates As RptRunDate = GetReportRunDate_MonthEnd(conn, tran)

                        Dim swapCurveDate As Date
                        'If  Month-end falls on Saturday, the swap curve of last working date prior month-end should be used.
                        If rptDates.RunDate.Value.DayOfWeek = DayOfWeek.Saturday Then
                            swapCurveDate = rptDates.PrevBusinessDate.Value
                        Else
                            swapCurveDate = rptDates.RunDate.Value
                        End If


                        config.SwapCurveFtpFile = FileUtils.ResolveFileName(config.SwapCurveFtpFile, swapCurveDate)
                        config.SwapCurveJPYFtpFile = FileUtils.ResolveFileName(config.SwapCurveJPYFtpFile, swapCurveDate)

                        Dim runDateIS9Prov As Date = rptDates.MonthendDate.Value.AddDays(1)
                        config.ProvisionFtpFile = FileUtils.ResolveFileName(config.ProvisionFtpFile, runDateIS9Prov)

                        If File.Exists(config.SwapCurveFtpFile) Then
                            File.Copy(config.SwapCurveFtpFile, config.SwapCurveDatFile, True)
                        Else
                            Utility.InsertExcp(strKPS_ref, "KPS to TFS yield curve interface file is not found.", conn, tran)
                            ' The program should end if the file is not found, according to the old source code.
                            Throw New Exception($"KPS to TFS yield curve interface file is not found! " & vbCrLf & config.SwapCurveFtpFile)
                        End If

                        If File.Exists(config.SwapCurveJPYFtpFile) Then
                            ' TODO: [REMOVE COMMENT AFTER REVIEW] Overwrite flag should be set here like the file copying of SwapCurveFtpFile and ProvisionFtpFile.
                            File.Copy(config.SwapCurveJPYFtpFile, config.SwapCurveJPYDatFile, True)
                        Else
                            Utility.InsertExcp(strTCL_ref, "TCL to TFS jpy yield curve interface file is not found.", conn, tran)

                            ' The program should end if the file is not found, according to the old source code.
                            Throw New Exception($"TCL to TFS jpy yield curve interface file is not found! " & vbCrLf & config.SwapCurveJPYFtpFile)

                        End If

                        Utility.GenRptSwapCurveCsv(strKPS_ref, swapCurveDate, config.SwapCurveDatFile, config.SwapCurveJPYDatFile, config.SwapCurveCsvFile, conn, tran)

                        If File.Exists(config.ProvisionFtpFile) Then
                            File.Copy(config.ProvisionFtpFile, config.ProvisionDatFile, True)
                            _logger.Info($"IS9 Provision interface file is found!")
                        Else
                            _logger.Info($"IS9 Provision interface file " & config.ProvisionFtpFile & " is not found!")
                        End If

                        If File.Exists(flagFile) Then
                            Dim workingDir As String = Path.GetDirectoryName(config.SwapCurveDatFile)
                            Dim ftpDir As String = Path.GetDirectoryName(config.SwapCurveFtpFile)

                            Dim ftpForfUploadCsv As String = ftpDir & "TFSTOTFS.RPT_FORF_UPLOAD.csv"
                            Dim destForfUploadCsv As String = workingDir & "RPT_FORF_UPLOAD.csv"

                            If File.Exists(ftpForfUploadCsv) Then
                                File.Copy(ftpForfUploadCsv, destForfUploadCsv, True)
                            End If

                            Dim ftpRepriceUploadCsv As String = ftpDir & "TFSTOTFS.RPT_REPRICE_UPLOAD.csv"
                            Dim destRepriceUploadCsv As String = workingDir & "RPT_REPRICE_UPLOAD.csv"

                            If File.Exists(ftpRepriceUploadCsv) Then
                                File.Copy(ftpRepriceUploadCsv, destRepriceUploadCsv, True)
                            End If

                            ' TODO: [REMOVE COMMENT AFTER REVIEW] Removed the postfix " rows=100".
                            Dim sqlLoaderFiles As New ReadOnlyCollection(Of ReadOnlyCollection(Of String))({
                                New ReadOnlyCollection(Of String)({
                                    config.RptForfUploadCtl,
                                    config.RptForfUploadLog
                                }),
                                New ReadOnlyCollection(Of String)({
                                    config.RptRepriceUploadCtl,
                                    config.RptRepriceUploadLog
                                }),
                                New ReadOnlyCollection(Of String)({
                                    config.RptSwapCurveCtl,
                                    config.RptSwapCurveLog
                                }),
                                New ReadOnlyCollection(Of String)({
                                    config.LoadIs9ProvisionCtl,
                                    config.LoadIs9ProvisionLog
                                })
                            })
                            For Each sqlLoader As ReadOnlyCollection(Of String) In sqlLoaderFiles
                                CallSqlLoader(dbConfig.RptDSN, dbConfig.RptUserID, dbConfig.RptPwd, workingDir, sqlLoader(0), sqlLoader(1))
                            Next

                            Utility.UpdateRptIS9Provision(conn, tran)

                            _logger.Info($"Delete flag file {flagFile}...")
                            File.Delete(flagFile)
                        Else
                            ' TODO: It is unclear whether the program MUST exit if the flagFile is not found.
                            Throw New Exception($"Flag file {flagFile} not found ...")
                        End If

                        ' TODO: These following steps occur after the flag file has been deleted. Refactored into separate Try...Catch block such that failure hereafter does not affect the exit code of this program.
                        Utility.ValidateForfaiting(conn, tran)
                        _logger.Info($"ValidateForfaiting Finished")
                        Utility.ValidateReprice(conn, tran)
                        _logger.Info($"ValidaterReprice Finished")
                        Utility.ValidateRepriceForOthers(conn, tran)
                        _logger.Info($"ValidaterRepriceForOthers Finished")
                        Utility.ValidateRepriceForPeriod(conn, tran)
                        _logger.Info($"ValidaterRepriceForPeriod Finished")
                        Utility.ValidateFullSet(conn, tran)
                        _logger.Info($"ValidateFullSet Finished")
                        Utility.ValidateCurve(conn, tran)
                        _logger.Info($"ValidateCurve Finished")
                        Utility.InsertForfaitingMaster(conn, tran)
                        _logger.Info($"InsertForfaitingMaster Finished")
                        Utility.UpdateSwapCurveStartDate(conn, tran)
                        _logger.Info($"UpdateSwapCurveStartDate Finished")

                        Utility.CheckUploadCurveIsNotExist(strKPS_ref, conn, tran)


#If DEBUG Then
                        tran.Rollback()
#Else
                        tran.Commit()
#End If
                    Catch ex As Exception
                        tran.Rollback()
                        Throw New Exception("Error occured during Forf_Upload, all transactions are rollbacked", ex)
                    End Try
                End Using
            End Using

            If (config.RunRepaySchFlag.ToUpper = "Y") Then
                _logger.Info($"Load RunGenRepaySch")
                ' TODO: [REMOVE COMMENT AFTER REVIEW] Created flag for Forf_GenRepaySch.
                File.Create(flagFileForfGenRepaySch).Close()
                Utility.RunProcess(config.RunRepaySchExe, config.RunRepaySchArgs)
                _logger.Info($"RunGenRepaySch Finished")
            End If

            ' Remove KPS files from sFTP folder
            Utility.RemoveKPSFiles(config.SwapCurveFtpFile)

            ' Remove IS9 Provision files from sFTP folder
            Utility.RemoveIS9Files(config.ProvisionFtpFile)

            _logger.Info("End Forf_Upload...")

        Catch ex As Exception
            'if no any special purpose/handle to the exception, throw all error here 
            _logger.Error($"Error occured in Forf_Upload : {vbNewLine}{GetExceptionMsgs(ex)}")
            Throw New Exception("Please look into inner exception", ex)
        End Try
    End Sub

    Public Function GetConfig(iniFile As String) As IniConfig
        Dim config As IniConfig = New IniConfig()
        Dim varINIFile As String = Path.GetFullPath(My.Settings.INI_FILE)

        _logger.Debug($"Config : {JsonConvert.SerializeObject(config)}")

        config.SwapCurveFtpFile = GetProfileStringVal("Forf_Upload", "SWAPCURVEFTPFILE", varINIFile)
        config.SwapCurveDatFile = GetProfileStringVal("Forf_Upload", "SWAPCURVEDATFILE", varINIFile)
        config.SwapCurveCsvFile = GetProfileStringVal("Forf_Upload", "SWAPCURVECSVFILE", varINIFile)
        config.ProvisionFtpFile = GetProfileStringVal("Forf_Upload", "PROVISIONFTPFILE", varINIFile)
        config.ProvisionDatFile = GetProfileStringVal("Forf_Upload", "PROVISIONDATFILE", varINIFile)
        config.RunRepaySchFlag = GetProfileStringVal("Forf_Upload", "RUNREPAYSCHFLAG", varINIFile)
        config.RunRepaySchExe = GetProfileStringVal("Forf_Upload", "GENREPAYSCH", varINIFile)
        config.RunRepaySchArgs = GetProfileStringVal("Forf_Upload", "GENREPAYSCHARGS", varINIFile)
        config.SwapCurveJPYDatFile = GetProfileStringVal("Forf_Upload", "SWAPCURVEJPYDATFILE", varINIFile)
        config.SwapCurveJPYFtpFile = GetProfileStringVal("Forf_Upload", "SWAPCURVEJPYFTPFILE", varINIFile)

        config.RptForfUploadCtl = config.SwapCurveJPYFtpFile = GetProfileStringVal("Forf_Upload", "RPTFORFUPLOADCTL", varINIFile)
        config.RptForfUploadLog = config.SwapCurveJPYFtpFile = GetProfileStringVal("Forf_Upload", "RPTFORFUPLOADLOG", varINIFile)
        config.RptRepriceUploadCtl = config.SwapCurveJPYFtpFile = GetProfileStringVal("Forf_Upload", "RPTREPRICEUPLOADCTL", varINIFile)
        config.RptRepriceUploadLog = config.SwapCurveJPYFtpFile = GetProfileStringVal("Forf_Upload", "RPTREPRICEUPLOADLOG", varINIFile)
        'config.RptCurveUsageCtl = config.SwapCurveJPYFtpFile = GetProfileStringVal("Forf_Upload", "RPTCURVEUSAGECTL", varINIFile)
        'config.RptCurveUsageLog = config.SwapCurveJPYFtpFile = GetProfileStringVal("Forf_Upload", "RPTCURVEUSAGELOG", varINIFile)
        config.RptSwapCurveCtl = config.SwapCurveJPYFtpFile = GetProfileStringVal("Forf_Upload", "RPTSWAPCURVECTL", varINIFile)
        config.RptSwapCurveLog = config.SwapCurveJPYFtpFile = GetProfileStringVal("Forf_Upload", "RPTSWAPCURVELOG", varINIFile)
        config.LoadIs9ProvisionCtl = config.SwapCurveJPYFtpFile = GetProfileStringVal("Forf_Upload", "LOADIS9PROVISIONCTL", varINIFile)
        config.LoadIs9ProvisionLog = config.SwapCurveJPYFtpFile = GetProfileStringVal("Forf_Upload", "LOADIS9PROVISIONLOG", varINIFile)

        _logger.Debug($"Config : {JsonConvert.SerializeObject(config)}")
        Return config
    End Function
End Module
