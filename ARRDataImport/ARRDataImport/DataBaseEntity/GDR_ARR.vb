Imports BEA.TFS.Common.ExtensionMethods

''' <summary>
''' Provide function to handle IO for GDR_ARR
''' </summary>
Module GDR_ARR
    ' No reason to declare the table structure here in this program, despite the insert Import function did...

    Public Sub ImportARRObject(ARRFileObject As ARRFile, conn As IDbConnection, tran As IDbTransaction)
        Using cmd As IDbCommand = conn.CreateCommand
            cmd.ExecuteNonQuery(
                AppConst.INSERT_INTO_ARR,
                New IDbDataParameter() {
                    cmd.CreateParameter("RATE_CODE", ARRFileObject.RATE_CODE),
                    cmd.CreateParameter("FCC", ARRFileObject.FCC),
                    cmd.CreateParameter("CONVENT_TYPE", ARRFileObject.CONVENT_TYPE),
                    cmd.CreateParameter("K_DAYS", ARRFileObject.K_DAYS),
                    cmd.CreateParameter("END_DATE", ARRFileObject.END_DATE),
                    cmd.CreateParameter("START_DATE", ARRFileObject.START_DATE),
                    cmd.CreateParameter("DAYS_OF_INT", ARRFileObject.DAYS_OF_INT),
                    cmd.CreateParameter("SIM_RATE_UNADJ", ARRFileObject.SIM_RATE_UNADJ),
                    cmd.CreateParameter("COMP_RATE_UNADJ", ARRFileObject.COMP_RATE_UNADJ),
                    cmd.CreateParameter("COMP_RATE_ADJ", ARRFileObject.COMP_RATE_ADJ),
                    cmd.CreateParameter("SIM_RATE_ADJ", ARRFileObject.SIM_RATE_ADJ),
                    cmd.CreateParameter("FLOOR_RATE", ARRFileObject.FLOOR_RATE),
                    cmd.CreateParameter("BAL_OUTSTD_HKD", ARRFileObject.FLOOR_IND),
                    cmd.CreateParameter("COMP_RATE_ADJ_CAS_ON", ARRFileObject.COMP_RATE_ADJ_CAS_ON),
                    cmd.CreateParameter("SIM_RATE_ADJ_CAS_ON", ARRFileObject.SIM_RATE_ADJ_CAS_ON),
                    cmd.CreateParameter("FLOOR_IND_CAS_ON", ARRFileObject.FLOOR_IND_CAS_ON),
                    cmd.CreateParameter("COMP_RATE_ADJ_CAS_1W", ARRFileObject.COMP_RATE_ADJ_CAS_1W),
                    cmd.CreateParameter("SIM_RATE_ADJ_CAS_1W", ARRFileObject.SIM_RATE_ADJ_CAS_1W),
                    cmd.CreateParameter("FLOOR_IND_CAS_1W", ARRFileObject.FLOOR_IND_CAS_1W),
                    cmd.CreateParameter("COMP_RATE_ADJ_CAS_1M", ARRFileObject.COMP_RATE_ADJ_CAS_1M),
                    cmd.CreateParameter("SIM_RATE_ADJ_CAS_1M", ARRFileObject.SIM_RATE_ADJ_CAS_1M),
                    cmd.CreateParameter("FLOOR_IND_CAS_1M", ARRFileObject.FLOOR_IND_CAS_1M),
                    cmd.CreateParameter("COMP_RATE_ADJ_CAS_2M", ARRFileObject.COMP_RATE_ADJ_CAS_2M),
                    cmd.CreateParameter("SIM_RATE_ADJ_CAS_2M", ARRFileObject.SIM_RATE_ADJ_CAS_2M),
                    cmd.CreateParameter("FLOOR_IND_CAS_2M", ARRFileObject.FLOOR_IND_CAS_2M),
                    cmd.CreateParameter("COMP_RATE_ADJ_CAS_3M", ARRFileObject.COMP_RATE_ADJ_CAS_3M),
                    cmd.CreateParameter("SIM_RATE_ADJ_CAS_3M", ARRFileObject.SIM_RATE_ADJ_CAS_3M),
                    cmd.CreateParameter("FLOOR_IND_CAS_3M", ARRFileObject.FLOOR_IND_CAS_3M),
                    cmd.CreateParameter("COMP_RATE_ADJ_CAS_6M", ARRFileObject.COMP_RATE_ADJ_CAS_6M),
                    cmd.CreateParameter("SIM_RATE_ADJ_CAS_6M", ARRFileObject.SIM_RATE_ADJ_CAS_6M),
                    cmd.CreateParameter("FLOOR_IND_CAS_6M", ARRFileObject.FLOOR_IND_CAS_6M),
                    cmd.CreateParameter("COMP_RATE_ADJ_CAS_12M", ARRFileObject.COMP_RATE_ADJ_CAS_12M),
                    cmd.CreateParameter("SIM_RATE_ADJ_CAS_12M", ARRFileObject.SIM_RATE_ADJ_CAS_12M),
                    cmd.CreateParameter("FLOOR_IND_CAS_12M", ARRFileObject.FLOOR_IND_CAS_12M),
                    cmd.CreateParameter("RUN_DATE", ARRFileObject.RUN_DATE)
                },
                tran
            )
        End Using
    End Sub
End Module
