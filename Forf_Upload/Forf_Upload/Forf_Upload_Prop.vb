Public Class Forf_Upload_Prop
    Public Class TB_RPT_FORF_MASTER
        Public Property BillsRef As String
        Public Property Tenor As String
        Public Property ModelType As String
        Public Property BillsCcy As String
        Public Property OsBalFcy As Decimal
        Public Property CREATE_DATE As Date
        Public Property DUE_DATE As Date
        Public Property INT_RATE As Decimal
    End Class
    Public Class TB_RPT_EXCPT
        Public Property BillsRef As String
        Public Property ErrorMessage As String
    End Class
    Public Class TB_RPT_REPRICE_UPLOAD
        Public Property BILLS_REF As String
        Public Property START_DATE As Date
        Public Property END_DATE As Date
    End Class
End Class
