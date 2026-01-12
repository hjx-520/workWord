Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Public Class DataSecurity
    Private Const Key As String = "BEAICO"
    Private Const IV As String = "BEAICO"


    Public Shared Function Encrypt(plainText As String) As String

        ' Password null check
        If plainText Is Nothing OrElse plainText.Length <= 0 Then
            Throw New ArgumentNullException("The input text is null")
        End If

        Dim encryptedPassword As String

        Using AES256 As New RijndaelManaged()

            ApplyAES256Settings(AES256)

            Using encryptStream As New MemoryStream()
                Using cryptStream As New CryptoStream(encryptStream, AES256.CreateEncryptor(), CryptoStreamMode.Write)
                    Using sw As New StreamWriter(cryptStream)
                        'Write all data to the stream
                        sw.Write(plainText)
                    End Using

                    encryptedPassword = Convert.ToBase64String(encryptStream.ToArray())
                End Using
            End Using
        End Using

        Return encryptedPassword
    End Function

    Public Shared Function Decrypt(cipherText As String) As String
        If cipherText Is Nothing OrElse cipherText.Length <= 0 Then
            Throw New ArgumentNullException("The input text is null")
        End If

        If Not cipherText.Length Mod 4 = 0 Then
            Throw New FormatException("The input password is not in a correct format. Please ensure the password is encrypted with AES256.")
        End If

        Dim decryptedPassword As String


        Using AES256 As New RijndaelManaged()
            ApplyAES256Settings(AES256)

            Using decryptStream As New MemoryStream(Convert.FromBase64String(cipherText))
                Using cryptStream As New CryptoStream(decryptStream, AES256.CreateDecryptor(), CryptoStreamMode.Read)
                    ' Convert the memory stream into string
                    Using reader As New StreamReader(cryptStream)
                        ' Read the decrypting stream
                        decryptedPassword = reader.ReadToEnd()
                    End Using
                End Using
            End Using
        End Using
        Return decryptedPassword
    End Function

    Private Shared Sub ApplyAES256Settings(ByRef aes As RijndaelManaged)

        ' Create salt through MD5 hashing using plainText as base
        Dim salt As Byte() = HashByMD5(Key)

        Dim rfc2898 As New Rfc2898DeriveBytes(Key, salt)

        ' Use AES256 encryption method
        aes.KeySize = 256
        aes.BlockSize = 128
        aes.Mode = CipherMode.CBC

        aes.Key = rfc2898.GetBytes(aes.KeySize / 8)
        aes.IV = rfc2898.GetBytes(aes.BlockSize / 8)
    End Sub

    Private Shared Function HashByMD5(source As String) As Byte()
        Dim md5Service As New MD5CryptoServiceProvider()
        Return md5Service.ComputeHash(Encoding.UTF8.GetBytes(source))
    End Function
End Class
