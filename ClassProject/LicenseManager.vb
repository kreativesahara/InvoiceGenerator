Imports System.IO
Imports System.Security.Cryptography
Imports System.Text

Friend NotInheritable Class LicenseManager
    Public Shared ReadOnly Property AppFolder As String
        Get
            Return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "InvoiceGenerator")
        End Get
    End Property

    Private Shared ReadOnly InstallFile As String = Path.Combine(AppFolder, "install.dat")
    Private Shared ReadOnly LicenseFilePath As String = Path.Combine(AppFolder, "license.lic")

    ' Replace this with your vendor/public RSA key XML (ToXmlString(False) style)
    Private Shared ReadOnly PublicKeyXml As String = "<RSAKeyValue>
<Modulus>wlHBheuDwPbMWJb25PzuDM0A1ecSHpGJHSVBKXSPX5bIUY6chjyEnPRuZlGLXVFSdU48T/ae19AMDw9kGePzFNS1UFpBelLgRJoZ1AaMvb+t/N/YBlQfflzKjslfMufgCzVh8nN3IboOLlBlVetmhU2Bt+KXe3PYVbRj8CiZn9X594zT9I3qwZ7nbn7w2eo7xu0VDrRAbQ2lhstDyFJrfqB56sBYjGx8lRPBhBwPpMjEJwEh26+JVJgueTsrgob4YUEa9iXFW46KngEqIpcWGYtekl8nt/NKjnf3DummK+kKfC9RYbhFToFS/5GtTezrpDFIVPAd8eYUOgsL4M8Q+WpFpBUL2gL6HgkPTVnvncP89pcN4OLDizlYoctJdws+3su5yax/N+YUnNBP+w2GDF2QowISYYvfzFr05zQw1l6vkmAxC0V5MVUlEqI/jjNtOwWz5O1zrXE3H4zq3TH8UnQDYeryzdzmbOXmav0aXjROS72VA+FgyotHKwZiZ+XMHJxHlHQCS08qtAlV5aY+kBVPd1/30Zxj8ShNJzmN6K6l+ZU5pxH9Szlq5FSZTWs17hAKt/2ASj4JiJoPcd4NyN+cpd0RfX1GSLy1TK0MrcblesvLyXWH8B2ZkhxWc9+yKQmtbh1vkTG6MBxt6TC2uCChhk58iFFAEB1nm0lRS6k=</Modulus>
<Exponent>AQAB</Exponent>
</RSAKeyValue>"

    Private Sub New()
    End Sub

    Public Shared Sub EnsureAppFolder()
        If Not Directory.Exists(AppFolder) Then Directory.CreateDirectory(AppFolder)
    End Sub

    Public Shared Function GetInstallDate() As DateTime
        If File.Exists(InstallFile) Then
            Dim s = File.ReadAllText(InstallFile).Trim()
            Dim dt As DateTime
            If DateTime.TryParse(s, dt) Then Return dt.ToUniversalTime()
        End If
        Dim nowUtc = DateTime.UtcNow
        File.WriteAllText(InstallFile, nowUtc.ToString("o"))
        Return nowUtc
    End Function

    Public Shared Function TrialDaysLeft() As Integer
        Dim install = GetInstallDate()
        Dim days = CInt((install.AddDays(60) - DateTime.UtcNow).TotalDays)
        If days < 0 Then Return 0
        Return days
    End Function

    Public Shared Function IsTrialActive() As Boolean
        Dim install = GetInstallDate()
        Return DateTime.UtcNow <= install.AddDays(60)
    End Function

    ' Validate license and extract expiry and client id from payload
    Public Shared Function TryValidateLicense(ByRef expiry As DateTime, ByRef licensedClientId As String) As Boolean
        expiry = DateTime.MinValue
        licensedClientId = String.Empty
        If Not File.Exists(LicenseFilePath) Then Return False
        Dim lines = File.ReadAllLines(LicenseFilePath)
        If lines.Length < 2 Then Return False
        Dim payload = lines(0).Trim()
        Dim signatureBase64 = lines(1).Trim()
        Dim signatureBytes As Byte()
        Try
            signatureBytes = Convert.FromBase64String(signatureBase64)
        Catch
            Return False
        End Try

        Dim payloadBytes = Encoding.UTF8.GetBytes(payload)
        Using rsa As New RSACryptoServiceProvider()
            Try
                rsa.PersistKeyInCsp = False
                rsa.FromXmlString(PublicKeyXml)
                Dim shaOid = CryptoConfig.MapNameToOID("SHA256")
                Dim verified = rsa.VerifyData(payloadBytes, shaOid, signatureBytes)
                If Not verified Then Return False
            Catch
                Return False
            End Try
        End Using

        ' payload expected format: expiry=ISO-8601;clientid=GUID
        For Each part In payload.Split(";"c)
            Dim kv = part.Split("="c)
            If kv.Length = 2 Then
                Dim key = kv(0).Trim().ToLower()
                Dim val = kv(1).Trim()
                If key = "expiry" Then
                    Dim dt As DateTime
                    If DateTime.TryParse(val, dt) Then expiry = dt.ToUniversalTime()
                ElseIf key = "clientid" Then
                    licensedClientId = val
                End If
            End If
        Next

        If expiry = DateTime.MinValue OrElse String.IsNullOrEmpty(licensedClientId) Then Return False
        Return True
    End Function

    Public Shared Function IsLicensed(localClientId As String, ByRef expiryUtc As DateTime) As Boolean
        expiryUtc = DateTime.MinValue
        Dim licClientId As String = String.Empty
        If Not TryValidateLicense(expiryUtc, licClientId) Then Return False
        If String.IsNullOrEmpty(localClientId) Then Return False
        If Not String.Equals(localClientId, licClientId, StringComparison.OrdinalIgnoreCase) Then Return False
        If DateTime.UtcNow > expiryUtc Then Return False
        Return True
    End Function
End Class
