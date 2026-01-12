Public Class DataSourceConfig
    Public Property TIDSN As String
    Public Property TIUserID As String
    Public Property TIPwd As String
    Public Property RptDSN As String
    Public Property RptUserID As String
    Public Property RptPwd As String
    Public Property TIGlobalDSN As String
    Public Property TIGlobalUserID As String
    Public Property TIGlobalPwd As String

    Public Overrides Function Equals(obj As Object) As Boolean
        Dim config = TryCast(obj, DataSourceConfig)
        Return config IsNot Nothing AndAlso
               TIDSN = config.TIDSN AndAlso
               TIUserID = config.TIUserID AndAlso
               TIPwd = config.TIPwd AndAlso
               RptDSN = config.RptDSN AndAlso
               RptUserID = config.RptUserID AndAlso
               RptPwd = config.RptPwd AndAlso
               TIGlobalDSN = config.TIGlobalDSN AndAlso
               TIGlobalUserID = config.TIGlobalUserID AndAlso
               TIGlobalPwd = config.TIGlobalPwd
    End Function
End Class
