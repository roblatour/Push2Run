Imports System.Reflection
Imports System.Security.Cryptography

Module modSecurity

    <Obfuscation(Feature:="Apply to member * when method: virtualization", Exclude:=False)>
    Public Class EncryptionClass

        Private Shared m_originalPassPhrase As String = "úÉ£Ky½Üå²ÄB†EQr•£ð9}]}³¶=®/[S}>[pí" 'same as m_strPassPhrase below
        Private Shared m_PassPhraseAddonForEncryption As String = String.Empty
        Private Shared m_PassPhraseAddonForDecryption As String = String.Empty

        Private Shared m_strPassPhrase As String = "úÉ£Ky½Üå²ÄB†EQr•£ð9}]}³¶=®/[S}>[pí" '---- any text string is good here
        Private Shared m_strHashAlgorithm As String = "MD5" '--- we are doing MD5 encryption - can be "SHA1"
        Private Shared m_strPasswordIterations As Integer = 100 '--- can be any number
        Private Shared m_strInitVector As String = "@1B2c3fia222W7H8" '--- must be 16 bytes
        Private Shared m_intKeySize As Integer = 256 '--- can be 192 or 128
        Private Shared m_strSaltValue As String = "s@aferlGHA2"

        '    'Encrypt Function: 

        <System.Diagnostics.DebuggerStepThrough()> Friend Shared Function Encrypt(ByVal PlainText As String) As String

            If PlainText = String.Empty Then Return String.Empty

            Dim strReturn As String = String.Empty

            ' Convert strings into byte arrays.
            ' Let us assume that strings only contain ASCII codes.
            ' If strings include Unicode characters, use Unicode, UTF7, or UTF8 
            ' encoding.

            Try
                Dim initVectorBytes As Byte()
                initVectorBytes = System.Text.Encoding.ASCII.GetBytes(m_strInitVector)

                Dim saltValueBytes As Byte()
                saltValueBytes = System.Text.Encoding.ASCII.GetBytes(m_strSaltValue)

                ' Convert our plaintext into a byte array.
                ' Let us assume that plaintext contains UTF8-encoded characters.
                Dim plainTextBytes As Byte()
                plainTextBytes = System.Text.Encoding.UTF8.GetBytes(PlainText)

                ' First, we must create a password, from which the key will be derived.
                ' This password will be generated from the specified passphrase and 
                ' salt value. The password will be created using the specified hash 
                ' algorithm. Password creation can be done in several iterations.

                Dim password As Rfc2898DeriveBytes

                password = New Rfc2898DeriveBytes(m_strPassPhrase & m_PassPhraseAddonForEncryption,
                saltValueBytes,
                m_strPasswordIterations)

                ' Use the password to generate pseudo-random bytes for the encryption
                ' key. Specify the size of the key in bytes (instead of bits).
                Dim keyBytes As Byte()
                Dim intKeySize As Integer = 0

                intKeySize = CType((m_intKeySize / 8), Integer)

                keyBytes = password.GetBytes(intKeySize)

                ' Create uninitialized Rijndael encryption object.
                Dim symmetricKey As System.Security.Cryptography.RijndaelManaged
                symmetricKey = New System.Security.Cryptography.RijndaelManaged

                ' It is reasonable to set encryption mode to Cipher Block Chaining
                ' (CBC). Use default options for other symmetric key parameters.
                symmetricKey.Mode = System.Security.Cryptography.CipherMode.CBC

                'symmetricKey.Padding = PaddingMode.Zeros


                ' Generate encryptor from the existing key bytes and initialization 
                ' vector. Key size will be defined based on the number of the key 
                ' bytes.
                Dim encryptor As System.Security.Cryptography.ICryptoTransform
                encryptor = symmetricKey.CreateEncryptor(keyBytes, initVectorBytes)

                ' Define memory stream which will be used to hold encrypted data.
                Dim memoryStream As System.IO.MemoryStream
                memoryStream = New System.IO.MemoryStream

                ' Define cryptographic stream (always use Write mode for encryption).
                Dim cryptoStream As System.Security.Cryptography.CryptoStream
                cryptoStream = New System.Security.Cryptography.CryptoStream(memoryStream,
                encryptor,
                System.Security.Cryptography.CryptoStreamMode.Write)
                ' Start encrypting.
                cryptoStream.Write(plainTextBytes, 0, plainTextBytes.Length)

                ' Finish encrypting.
                cryptoStream.FlushFinalBlock()

                ' Convert our encrypted data from a memory stream into a byte array.
                Dim cipherTextBytes As Byte()
                cipherTextBytes = memoryStream.ToArray()

                ' Close both streams.
                memoryStream.Close()
                cryptoStream.Close()

                ' Convert encrypted data into a base64-encoded string.
                Dim cipherText As String
                cipherText = Convert.ToBase64String(cipherTextBytes)

                ' Return encrypted string.
                strReturn = cipherText

            Catch ex As Exception
                strReturn = String.Empty
            End Try

            Return strReturn

        End Function

        'Decrypt Function: 

        <Obfuscation(Feature:="Apply to member * when method: virtualization", Exclude:=False)>
        <System.Diagnostics.DebuggerStepThrough()> Friend Shared Function Decrypt(ByVal CipherText As String) As String

            If CipherText = String.Empty Then Return String.Empty

            Dim strReturn As String = String.Empty

            ' Convert strings defining encryption key characteristics into byte
            ' arrays. Let us assume that strings only contain ASCII codes.
            ' If strings include Unicode characters, use Unicode, UTF7, or UTF8
            ' encoding.

            Try

                Dim initVectorBytes As Byte()
                initVectorBytes = System.Text.Encoding.ASCII.GetBytes(m_strInitVector)

                Dim saltValueBytes As Byte()
                saltValueBytes = System.Text.Encoding.ASCII.GetBytes(m_strSaltValue)

                ' Convert our ciphertext into a byte array.
                Dim cipherTextBytes As Byte()
                cipherTextBytes = Convert.FromBase64String(CipherText)

                ' First, we must create a password, from which the key will be 
                ' derived. This password will be generated from the specified 
                ' passphrase and salt value. The password will be created using
                ' the specified hash algorithm. Password creation can be done in
                ' several iterations.

                Dim password As Rfc2898DeriveBytes

                password = New Rfc2898DeriveBytes(m_strPassPhrase & m_PassPhraseAddonForDecryption,
                saltValueBytes,
                m_strPasswordIterations)

                ' Use the password to generate pseudo-random bytes for the encryption
                ' key. Specify the size of the key in bytes (instead of bits).
                Dim keyBytes As Byte()
                Dim intKeySize As Integer

                intKeySize = CType((m_intKeySize / 8), Integer)

                keyBytes = password.GetBytes(intKeySize)

                ' Create uninitialized Rijndael encryption object.
                Dim symmetricKey As System.Security.Cryptography.RijndaelManaged
                symmetricKey = New System.Security.Cryptography.RijndaelManaged

                ' It is reasonable to set encryption mode to Cipher Block Chaining
                ' (CBC). Use default options for other symmetric key parameters.
                symmetricKey.Mode = System.Security.Cryptography.CipherMode.CBC

                'symmetricKey.Padding = PaddingMode.Zeros

                ' Generate decryptor from the existing key bytes and initialization 
                ' vector. Key size will be defined based on the number of the key 
                ' bytes.
                Dim decryptor As System.Security.Cryptography.ICryptoTransform
                decryptor = symmetricKey.CreateDecryptor(keyBytes, initVectorBytes)

                ' Define memory stream which will be used to hold encrypted data.
                Dim memoryStream As System.IO.MemoryStream
                memoryStream = New System.IO.MemoryStream(cipherTextBytes)

                ' Define memory stream which will be used to hold encrypted data.
                Dim cryptoStream As System.Security.Cryptography.CryptoStream
                cryptoStream = New System.Security.Cryptography.CryptoStream(memoryStream,
                decryptor,
                System.Security.Cryptography.CryptoStreamMode.Read)

                ' Since at this point we don't know what the size of decrypted data
                ' will be, allocate the buffer long enough to hold ciphertext;
                ' plaintext is never longer than ciphertext.
                Dim plainTextBytes As Byte()
                ReDim plainTextBytes(cipherTextBytes.Length)

                ' Start decrypting.
                Dim decryptedByteCount As Integer
                decryptedByteCount = cryptoStream.Read(plainTextBytes,
                0,
                plainTextBytes.Length)

                ' Close both streams.
                memoryStream.Close()
                cryptoStream.Close()

                ' Convert decrypted data into a string. 
                ' Let us assume that the original plaintext string was UTF8-encoded.
                Dim plainText As String
                plainText = System.Text.Encoding.UTF8.GetString(plainTextBytes,
                0,
                decryptedByteCount)

                ' Return decrypted string.
                strReturn = plainText

            Catch ex As Exception
                strReturn = String.Empty
            End Try

            Return strReturn

        End Function

        <Obfuscation(Feature:="Apply to member * when method: virtualization", Exclude:=False)>
        <System.Diagnostics.DebuggerStepThrough()> Friend Shared Sub ResetEncryptionPassPhrase()
            m_PassPhraseAddonForEncryption = String.Empty
        End Sub

        <Obfuscation(Feature:="Apply to member * when method: virtualization", Exclude:=False)>
        <System.Diagnostics.DebuggerStepThrough()> Friend Shared Sub ResetDecryptionPassPhrase()
            m_PassPhraseAddonForDecryption = String.Empty
        End Sub

        <Obfuscation(Feature:="Apply to member * when method: virtualization", Exclude:=False)>
        <System.Diagnostics.DebuggerStepThrough()> Friend Shared Sub UpdateEncryptionPassPhrase(ByVal PassPhraseAddOn)
            m_PassPhraseAddonForEncryption &= PassPhraseAddOn
        End Sub

        <Obfuscation(Feature:="Apply to member * when method: virtualization", Exclude:=False)>
        <System.Diagnostics.DebuggerStepThrough()> Friend Shared Sub UpdateDecryptionPassPhrase(ByVal PassPhraseAddOn)
            m_PassPhraseAddonForDecryption &= PassPhraseAddOn
        End Sub

    End Class

End Module

