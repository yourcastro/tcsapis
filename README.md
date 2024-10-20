Imports System.IO
Imports System.Security.Cryptography
Imports System.Text

Public Class ClsPDScoreCardEncryption
    Shared Function Encrypt(ByVal key As String, ByVal data As String) As String
        Dim encData As String = Nothing
        Dim keys As Byte()() = GetHashKeys(key)
        encData = EncryptString_Aes(data, keys(0), keys(1))

        Return encData
    End Function

    Shared Function Decrypt(ByVal key As String, ByVal data As String) As String
        Dim decData As String = Nothing
        Dim keys As Byte()() = GetHashKeys(key)

        decData = DecryptString_Aes(data, keys(0), keys(1))

        Return decData
    End Function

    Shared Function GetHashKeys(ByVal key As String) As Byte()()
        Dim result As Byte()() = New Byte(1)() {}
        Dim enc As Encoding = Encoding.UTF8
        Dim sha2 As SHA256 = New SHA256CryptoServiceProvider()
        Dim rawKey As Byte() = enc.GetBytes(key)
        Dim rawIV As Byte() = enc.GetBytes(key)
        Dim hashKey As Byte() = sha2.ComputeHash(rawKey)
        Dim hashIV As Byte() = sha2.ComputeHash(rawIV)
        Array.Resize(hashIV, 16)
        result(0) = hashKey
        result(1) = hashIV
        Return result
    End Function

    Shared Function EncryptString_Aes(ByVal plainText As String, ByVal Key() As Byte, ByVal IV() As Byte) As String
        ' Check arguments.
        If plainText Is Nothing OrElse plainText.Length <= 0 Then
            Throw New ArgumentNullException("plainText")
        End If
        If Key Is Nothing OrElse Key.Length <= 0 Then
            Throw New ArgumentNullException("Key")
        End If
        If IV Is Nothing OrElse IV.Length <= 0 Then
            Throw New ArgumentNullException("IV")
        End If
        Dim encrypted() As Byte

        ' Create an AesManaged object
        ' with the specified key and IV.
        Using aesAlg As New AesManaged()
            aesAlg.Padding = PaddingMode.None
            aesAlg.Key = Key
            aesAlg.IV = IV

            ' Create an encryptor to perform the stream transform.
            Dim encryptor As ICryptoTransform = aesAlg.CreateEncryptor(aesAlg.Key, aesAlg.IV)

            Using msEncrypt As New MemoryStream()
                Using csEncrypt As New CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write)
                    Using swEncrypt As New StreamWriter(csEncrypt)

                        swEncrypt.Write(plainText)
                    End Using
                    encrypted = msEncrypt.ToArray()
                End Using
            End Using
        End Using

        Return Convert.ToBase64String(encrypted)

    End Function 'EncryptString_Aes

    Shared Function DecryptString_Aes(ByVal cipherTextString As String, ByVal Key() As Byte, ByVal IV() As Byte) As String
        ' Check arguments.

        Dim cipherText() As Byte = Nothing
        cipherText = Convert.FromBase64String(cipherTextString)

        If cipherText Is Nothing OrElse cipherText.Length <= 0 Then
            Throw New ArgumentNullException("cipherText")
        End If
        If Key Is Nothing OrElse Key.Length <= 0 Then
            Throw New ArgumentNullException("Key")
        End If
        If IV Is Nothing OrElse IV.Length <= 0 Then
            Throw New ArgumentNullException("IV")
        End If
        ' Declare the string used to hold
        ' the decrypted text.
        Dim plaintext As String = Nothing

        ' Create an AesManaged object
        ' with the specified key and IV.
        Using aesAlg As New AesManaged
            aesAlg.Padding = PaddingMode.Zeros
            aesAlg.Key = Key
            aesAlg.IV = IV

            Dim decryptor As ICryptoTransform = aesAlg.CreateDecryptor(aesAlg.Key, aesAlg.IV)

            Using msDecrypt As New MemoryStream(cipherText)

                Using csDecrypt As New CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read)

                    Using srDecrypt As New StreamReader(csDecrypt)
                        plaintext = srDecrypt.ReadToEnd()
                    End Using
                End Using
            End Using
        End Using

        Return RegularExpressions.Regex.Replace(plaintext, "[^a-zA-Z0-9\\:_\- ]", "")

    End Function 'DecryptString_Aes 

End Class
