''' <summary>
''' Defines the structure for file ARRFile
''' </summary>
Public Class ARRFile
    Public Property RATE_CODE As Decimal
    Public Property FCC As Decimal
    Public Property CONVENT_TYPE As String
    Public Property K_DAYS As String
    Public Property END_DATE As String
    Public Property START_DATE As String
    Public Property DAYS_OF_INT As String
    Public Property COMP_RATE_UNADJ As String
    Public Property SIM_RATE_UNADJ As Decimal
    Public Property COMP_RATE_ADJ As Decimal
    Public Property SIM_RATE_ADJ As Decimal
    Public Property FLOOR_RATE As Decimal
    Public Property FLOOR_IND As String
    Public Property COMP_RATE_ADJ_CAS_ON As Decimal
    Public Property SIM_RATE_ADJ_CAS_ON As Decimal
    Public Property FLOOR_IND_CAS_ON As Decimal
    Public Property COMP_RATE_ADJ_CAS_1W As Decimal
    Public Property SIM_RATE_ADJ_CAS_1W As Decimal
    Public Property FLOOR_IND_CAS_1W As Decimal
    Public Property COMP_RATE_ADJ_CAS_1M As Decimal
    Public Property SIM_RATE_ADJ_CAS_1M As Decimal
    Public Property FLOOR_IND_CAS_1M As String
    Public Property COMP_RATE_ADJ_CAS_2M As Decimal
    Public Property SIM_RATE_ADJ_CAS_2M As Decimal
    Public Property FLOOR_IND_CAS_2M As String
    Public Property COMP_RATE_ADJ_CAS_3M As Decimal
    Public Property SIM_RATE_ADJ_CAS_3M As Decimal
    Public Property FLOOR_IND_CAS_3M As String
    Public Property COMP_RATE_ADJ_CAS_6M As Decimal
    Public Property SIM_RATE_ADJ_CAS_6M As Decimal
    Public Property FLOOR_IND_CAS_6M As String
    Public Property COMP_RATE_ADJ_CAS_12M As Decimal
    Public Property SIM_RATE_ADJ_CAS_12M As Decimal
    Public Property FLOOR_IND_CAS_12M As String
    Public Property RUN_DATE As Date



    ''' <summary>
    ''' Decode a line read from ARRFile into different fields.
    ''' </summary>
    ''' <param name="line">One single body line read from file ARRFile</param>
    ''' <returns></returns>
    Public Shared Function ReadLine(line As String) As ARRFile
        Dim ARRFileObject As New ARRFile()

        ' Run_Date and LastCalendarDate is not read from the line
        ARRFileObject.RATE_CODE = line.Substring(6, 8).Trim
        ARRFileObject.FCC = line.Substring(14, 3).Trim
        ARRFileObject.CONVENT_TYPE = line.Substring(17, 5).Trim
        ARRFileObject.K_DAYS = line.Substring(22, 3).Trim
        ARRFileObject.END_DATE = Convert.ToDateTime(line.Substring(25, 10).Trim)
        ARRFileObject.START_DATE = Convert.ToDateTime(line.Substring(35, 10).Trim)
        ARRFileObject.DAYS_OF_INT = If(line.Substring(45, 5).Trim = "", 0, Convert.ToInt16(line.Substring(45, 5).Trim))
        ARRFileObject.COMP_RATE_UNADJ = If(line.Substring(50, 11).Trim = "", 0D, Convert.ToDecimal(line.Substring(50, 11).Trim))
        ARRFileObject.SIM_RATE_UNADJ = If(line.Substring(61, 11).Trim = "", 0D, Convert.ToDecimal(line.Substring(61, 11).Trim))
        ARRFileObject.COMP_RATE_ADJ = If(line.Substring(72, 11).Trim = "", 0D, Convert.ToDecimal(line.Substring(72, 11).Trim))
        ARRFileObject.SIM_RATE_ADJ = If(line.Substring(83, 11).Trim = "", 0D, Convert.ToDecimal(line.Substring(83, 11).Trim))
        ARRFileObject.FLOOR_RATE = If(line.Substring(94, 11).Trim = "", 0D, Convert.ToDecimal(line.Substring(94, 11).Trim))
        ARRFileObject.FLOOR_IND = line.Substring(105, 1).Trim
        ARRFileObject.COMP_RATE_ADJ_CAS_ON = If(line.Substring(106, 11).Trim = "", 0D, Convert.ToDecimal(line.Substring(106, 11).Trim))
        ARRFileObject.SIM_RATE_ADJ_CAS_ON = If(line.Substring(117, 11).Trim = "", 0D, Convert.ToDecimal(line.Substring(117, 11).Trim))
        ARRFileObject.FLOOR_IND_CAS_ON = line.Substring(128, 1).Trim
        ARRFileObject.COMP_RATE_ADJ_CAS_1W = If(line.Substring(129, 11).Trim = "", 0D, Convert.ToDecimal(line.Substring(129, 11).Trim))
        ARRFileObject.SIM_RATE_ADJ_CAS_1W = If(line.Substring(140, 11).Trim = "", 0D, Convert.ToDecimal(line.Substring(140, 11).Trim))
        ARRFileObject.FLOOR_IND_CAS_1W = line.Substring(151, 1).Trim
        ARRFileObject.COMP_RATE_ADJ_CAS_1M = If(line.Substring(152, 11).Trim = "", 0D, Convert.ToDecimal(line.Substring(152, 11).Trim))
        ARRFileObject.SIM_RATE_ADJ_CAS_1M = If(line.Substring(163, 11).Trim = "", 0D, Convert.ToDecimal(line.Substring(163, 11).Trim))
        ARRFileObject.FLOOR_IND_CAS_1M = line.Substring(174, 1).Trim
        ARRFileObject.COMP_RATE_ADJ_CAS_2M = If(line.Substring(175, 11).Trim = "", 0D, Convert.ToDecimal(line.Substring(175, 11).Trim))
        ARRFileObject.SIM_RATE_ADJ_CAS_2M = If(line.Substring(186, 11).Trim = "", 0D, Convert.ToDecimal(line.Substring(186, 11).Trim))
        ARRFileObject.FLOOR_IND_CAS_2M = line.Substring(197, 1).Trim
        ARRFileObject.COMP_RATE_ADJ_CAS_3M = If(line.Substring(198, 11).Trim = "", 0D, Convert.ToDecimal(line.Substring(198, 11).Trim))
        ARRFileObject.SIM_RATE_ADJ_CAS_3M = If(line.Substring(209, 11).Trim = "", 0D, Convert.ToDecimal(line.Substring(209, 11).Trim))
        ARRFileObject.FLOOR_IND_CAS_3M = line.Substring(220, 1).Trim
        ARRFileObject.COMP_RATE_ADJ_CAS_6M = If(line.Substring(221, 11).Trim = "", 0D, Convert.ToDecimal(line.Substring(221, 11).Trim))
        ARRFileObject.SIM_RATE_ADJ_CAS_6M = If(line.Substring(232, 11).Trim = "", 0D, Convert.ToDecimal(line.Substring(232, 11).Trim))
        ARRFileObject.FLOOR_IND_CAS_6M = line.Substring(243, 1).Trim
        ARRFileObject.COMP_RATE_ADJ_CAS_12M = If(line.Substring(244, 11).Trim = "", 0D, Convert.ToDecimal(line.Substring(244, 11).Trim))
        ARRFileObject.SIM_RATE_ADJ_CAS_12M = If(line.Substring(255, 11).Trim = "", 0D, Convert.ToDecimal(line.Substring(255, 11).Trim))
        ARRFileObject.FLOOR_IND_CAS_12M = line.Substring(266, 1).Trim

        Return ARRFileObject
    End Function
End Class
